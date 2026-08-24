import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {
  IListConfiguration
} from '../models/IListConfiguration';

import {
  ISharePointItemsResponse,
  ISharePointSubmissionItem
} from '../models/ISharePointSubmissionItem';

import {
  IQualificationResult
} from '../models/IQualificationResult';

import {
  ISupplierSubmission
} from '../models/ISupplierSubmission';

import {
  ISupplierSubmissionService
} from './ISupplierSubmissionService';

import {
  SupplierQualificationUtility
} from '../utilities/SupplierQualificationUtility';

export class SupplierSubmissionService
  implements ISupplierSubmissionService {

  private readonly context: WebPartContext;
  private readonly webAbsoluteUrl: string;

  public constructor(
    context: WebPartContext
  ) {
    this.context = context;

    this.webAbsoluteUrl =
      context.pageContext.web.absoluteUrl.replace(
        /\/$/,
        ''
      );
  }

  /**
   * Retrieves submissions from all configured ESG questionnaire
   * lists and combines the results.
   *
   * The thresholds argument is retained for compatibility with the
   * existing ISupplierSubmissionService interface. The qualification
   * result is determined using tier-specific weighted thresholds.
   */
  public async getAllSubmissions(
  listConfigurations:
    readonly IListConfiguration[]
): Promise<ISupplierSubmission[]> {
    const listPromises:
      Array<Promise<ISupplierSubmission[]>> =
      listConfigurations.map(
        (
          configuration: IListConfiguration
        ): Promise<ISupplierSubmission[]> => {
          return this.getSubmissionsFromList(
            configuration
          );
        }
      );

    const groupedResults:
      ISupplierSubmission[][] =
      await Promise.all(listPromises);

    let combinedResults:
      ISupplierSubmission[] = [];

    groupedResults.forEach(
      (
        group: ISupplierSubmission[]
      ): void => {
        combinedResults =
          combinedResults.concat(group);
      }
    );

    combinedResults.sort(
      (
        first: ISupplierSubmission,
        second: ISupplierSubmission
      ): number => {
        return (
          this.getDateSortValue(second.created) -
          this.getDateSortValue(first.created)
        );
      }
    );

    return combinedResults;
  }

  /**
   * Retrieves all items from one configured SharePoint list.
   */
  private async getSubmissionsFromList(
    configuration:
      IListConfiguration
  ): Promise<ISupplierSubmission[]> {
    const encodedListTitle: string =
      this.escapeODataString(
        configuration.listTitle
      );

    const selectFields: string[] =
      this.getUniqueFieldNames([
        'Id',

        configuration
          .supplierNameInternalName,

        configuration
          .emailInternalName,

        configuration
          .contactNameInternalName,

        configuration
          .overallPercentageInternalName,

        'Created',
        'Modified'
      ]);

    const endpoint: string =
      `${this.webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle(` +
      `'${encodedListTitle}')/items` +
      `?$select=${selectFields.join(',')}` +
      '&$orderby=Created desc' +
      '&$top=5000';

    const rawItems:
      ISharePointSubmissionItem[] =
      await this.getAllPages(
        endpoint,
        configuration.listTitle
      );

    return rawItems.map(
      (
        item:
          ISharePointSubmissionItem
      ): ISupplierSubmission => {
        return this.mapSubmission(
          item,
          configuration
        );
      }
    );
  }

  /**
   * Retrieves all SharePoint REST result pages.
   */
  private async getAllPages(
    initialEndpoint: string,
    listTitle: string
  ): Promise<ISharePointSubmissionItem[]> {
    let allItems:
      ISharePointSubmissionItem[] = [];

    let nextEndpoint:
      string | undefined =
      initialEndpoint;

    while (nextEndpoint) {
      const response:
        SPHttpClientResponse =
        await this.context
          .spHttpClient
          .get(
            nextEndpoint,
            SPHttpClient
              .configurations
              .v1,
            {
              headers: {
                Accept:
                  'application/json;' +
                  'odata.metadata=none'
              }
            }
          );

      if (!response.ok) {
        const responseText:
          string =
          await response.text();

        throw new Error(
          `Unable to retrieve submissions from ` +
          `"${listTitle}". ` +
          `HTTP status: ${response.status} ` +
          `${response.statusText}. ` +
          `Response: ${responseText}`
        );
      }

      const responseData:
        ISharePointItemsResponse =
        await response.json() as
        ISharePointItemsResponse;

      const pageItems:
        ISharePointSubmissionItem[] =
        responseData.value || [];

      allItems =
        allItems.concat(
          pageItems
        );

      nextEndpoint =
        responseData[
        '@odata.nextLink'
        ] ||
        responseData[
        'odata.nextLink'
        ];
    }

    return allItems;
  }

  /**
   * Maps a SharePoint list item to the common supplier
   * submission model.
   *
   * OverallQuestionsPercentage is read exactly from SharePoint.
   * No score normalization or weighted value conversion occurs.
   */
  private mapSubmission(
    item:
      ISharePointSubmissionItem,

    configuration:
      IListConfiguration
  ): ISupplierSubmission {
    const supplierFieldValue:
      string | number | undefined =
      item[
      configuration
        .supplierNameInternalName
      ];

    const emailFieldValue:
      string | number | undefined =
      item[
      configuration
        .emailInternalName
      ];

    const contactNameFieldValue:
      string | number | undefined =
      item[
      configuration
        .contactNameInternalName
      ];

    const sharePointScoreValue:
      string | number | undefined =
      item[
      configuration
        .overallPercentageInternalName
      ];

    const supplierName: string =
      this.getTextValue(
        supplierFieldValue
      );

    const email: string =
      this.getTextValue(
        emailFieldValue
      );

    const submittedByName: string =
      this.getTextValue(
        contactNameFieldValue
      );

    /**
     * Read the exact numeric value returned by SharePoint.
     *
     * Examples:
     * "10%"   becomes 10
     * "9.99%" becomes 9.99
     * "5%"    becomes 5
     * "4.99%" becomes 4.99
     * "2%"    becomes 2
     * "1.99%" becomes 1.99
     *
     * Removing a display percent symbol is parsing only.
     * The numeric scale is not changed.
     */
    const overallPercentage: number =
  SupplierQualificationUtility
    .parseSharePointValue(
      sharePointScoreValue
    );

    const qualificationResult:
      IQualificationResult =
      this.evaluateExactSharePointScore(
        overallPercentage,
        configuration
      );

    return {
      key:
        `${configuration.tier}-` +
        `${item.Id.toString()}`,

      id:
        item.Id,

      sourceListTitle:
        configuration.listTitle,

      sourceItemUrl:
        this.buildItemUrl(
          configuration.listTitle,
          item.Id
        ),

      tier:
        configuration.tier,

      supplierName:
        supplierName ||
        'Supplier name unavailable',

      email,

      submittedByName,

      /**
       * This is the exact numeric value obtained from
       * OverallQuestionsPercentage.
       */
      overallPercentage,

      qualification:
        qualificationResult
          .qualification,

      riskRating:
        qualificationResult
          .riskRating,

      recommendation:
        qualificationResult
          .recommendation,

      created:
        item.Created,

      modified:
        item.Modified
    };
  }

  /**
   * Evaluates the exact SharePoint score against the
   * tier-specific thresholds.
   *
   * Tier 1:
   * Qualified >= 10
   * Conditional >= 9.99 and < 10
   * Not Qualified < 9.99
   *
   * Tier 2:
   * Qualified >= 5
   * Conditional >= 4.99 and < 5
   * Not Qualified < 4.99
   *
   * Tier 3:
   * Qualified >= 2
   * Conditional >= 1.99 and < 2
   * Not Qualified < 1.99
   */
  private evaluateExactSharePointScore(
    sharePointScore: number,
    configuration:
      IListConfiguration
  ): IQualificationResult {
    if (
      !isFinite(
        sharePointScore
      ) ||
      sharePointScore < 0
    ) {
      return {
        qualification:
          'Not Qualified',

        riskRating:
          'High Risk',

        recommendation:
          'Requires Review'
      };
    }

    if (
      sharePointScore >=
      configuration
        .qualifiedWeightedMinimum
    ) {
      return {
        qualification:
          'Qualified',

        riskRating:
          'Low Risk',

        recommendation:
          'Approve'
      };
    }

    if (
      sharePointScore >=
      configuration
        .conditionalWeightedMinimum
    ) {
      return {
        qualification:
          'Conditionally Qualified',

        riskRating:
          'Medium Risk',

        recommendation:
          'Corrective Action'
      };
    }

    return {
      qualification:
        'Not Qualified',

      riskRating:
        'High Risk',

      recommendation:
        'Requires Review'
    };
  }

  /**
   * Parses a SharePoint numeric or formatted percentage value.
   *
   * The method does not normalize, divide, multiply, or
   * otherwise change the numeric scale.
   */

  private getTextValue(
    value:
      string | number | undefined
  ): string {
    if (
      typeof value === 'string'
    ) {
      return value.trim();
    }

    if (
      typeof value === 'number'
    ) {
      return value.toString();
    }

    return '';
  }

  private getUniqueFieldNames(
    fieldNames:
      readonly string[]
  ): string[] {
    const uniqueFieldNames:
      string[] = [];

    fieldNames.forEach(
      (
        fieldName:
          string
      ): void => {
        if (
          fieldName &&
          uniqueFieldNames.indexOf(
            fieldName
          ) === -1
        ) {
          uniqueFieldNames.push(
            fieldName
          );
        }
      }
    );

    return uniqueFieldNames;
  }

  private getDateSortValue(
    dateValue: string
  ): number {
    const timeValue: number =
      new Date(
        dateValue
      ).getTime();

    return isNaN(
      timeValue
    )
      ? 0
      : timeValue;
  }

  private buildItemUrl(
    listTitle: string,
    itemId: number
  ): string {
    return (
      `${this.webAbsoluteUrl}/Lists/` +
      `${encodeURIComponent(listTitle)}` +
      `/DispForm.aspx?ID=` +
      `${itemId.toString()}`
    );
  }

  private escapeODataString(
    value: string
  ): string {
    return value.replace(
      /'/g,
      "''"
    );
  }
}