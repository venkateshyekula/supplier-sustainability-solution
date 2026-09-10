import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IListConfiguration } from '../models/IListConfiguration';

import {
  ISharePointItemsResponse,
  ISharePointSubmissionItem
} from '../models/ISharePointSubmissionItem';

import { IQualificationResult } from '../models/IQualificationResult';

import { ISupplierSubmission } from '../models/ISupplierSubmission';

import { ISupplierSubmissionService } from './ISupplierSubmissionService';

import {
  SupplierQualificationUtility
} from '../utilities/SupplierQualificationUtility';

export class SupplierSubmissionService
  implements ISupplierSubmissionService {

  private readonly context: WebPartContext;
  private readonly webAbsoluteUrl: string;

  public constructor(context: WebPartContext) {
    this.context = context;

    this.webAbsoluteUrl =
      context.pageContext.web.absoluteUrl.replace(
        /\/+$/,
        ''
      );
  }

  /**
   * Retrieves supplier submissions from all configured
   * questionnaire lists and combines the results.
   */
  public async getAllSubmissions(
    listConfigurations: readonly IListConfiguration[]
  ): Promise<ISupplierSubmission[]> {
    const listPromises: Array<Promise<ISupplierSubmission[]>> =
      listConfigurations.map(
        (
          configuration: IListConfiguration
        ): Promise<ISupplierSubmission[]> => {
          return this.getSubmissionsFromList(
            configuration
          );
        }
      );

    const groupedResults: ISupplierSubmission[][] =
      await Promise.all(
        listPromises
      );

    let combinedResults: ISupplierSubmission[] = [];

    groupedResults.forEach(
      (
        group: ISupplierSubmission[]
      ): void => {
        combinedResults =
          combinedResults.concat(
            group
          );
      }
    );

    combinedResults.sort(
      (
        first: ISupplierSubmission,
        second: ISupplierSubmission
      ): number => {
        return (
          this.getDateSortValue(
            second.created
          ) -
          this.getDateSortValue(
            first.created
          )
        );
      }
    );

    return combinedResults;
  }

  /**
   * Retrieves all items from one configured
   * SharePoint questionnaire list.
   */
  private async getSubmissionsFromList(
    configuration: IListConfiguration
  ): Promise<ISupplierSubmission[]> {
    const escapedListTitle: string =
      this.escapeODataString(
        configuration.listTitle
      );

    const encodedListTitle: string =
      encodeURIComponent(
        escapedListTitle
      );

    const selectFields: string[] =
      this.getUniqueFieldNames([
        'Id',
        configuration.supplierNameInternalName,
        configuration.emailInternalName,
        configuration.contactNameInternalName,
        configuration.overallPercentageInternalName,
        configuration.weightingInternalName,
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

    const rawItems: ISharePointSubmissionItem[] =
      await this.getAllPages(
        endpoint,
        configuration.listTitle
      );

    return rawItems.map(
      (
        item: ISharePointSubmissionItem
      ): ISupplierSubmission => {
        return this.mapSubmission(
          item,
          configuration
        );
      }
    );
  }

  /**
   * Retrieves every SharePoint REST result page.
   */
  private async getAllPages(
    initialEndpoint: string,
    listTitle: string
  ): Promise<ISharePointSubmissionItem[]> {
    let allItems: ISharePointSubmissionItem[] = [];

    let nextEndpoint: string | undefined =
      initialEndpoint;

    while (nextEndpoint) {
      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          nextEndpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept:
                'application/json;' +
                'odata.metadata=none'
            }
          }
        );

      if (!response.ok) {
        const responseText: string =
          await response.text();

        throw new Error(
          `Unable to retrieve submissions from ` +
          `"${listTitle}". ` +
          `HTTP status: ${response.status} ` +
          `${response.statusText}. ` +
          `Response: ${responseText}`
        );
      }

      const responseData: ISharePointItemsResponse =
        await response.json() as ISharePointItemsResponse;

      const pageItems: ISharePointSubmissionItem[] =
        responseData.value || [];

      allItems =
        allItems.concat(
          pageItems
        );

      nextEndpoint =
        responseData['@odata.nextLink'] ||
        responseData['odata.nextLink'];
    }

    return allItems;
  }

  /**
   * Maps the source SharePoint item to the common
   * supplier submission model.
   *
   * The Sustainability field is mapped to
   * overallPercentage for display.
   *
   * The Weighting field is mapped to weightedScore
   * and is used for qualification.
   *
   * No normalization, division, multiplication,
   * or percentage conversion is applied.
   */
  private mapSubmission(
    item: ISharePointSubmissionItem,
    configuration: IListConfiguration
  ): ISupplierSubmission {
    const supplierFieldValue:
      string | number | undefined =
      this.getStringOrNumberValue(
        item[
          configuration.supplierNameInternalName
        ]
      );

    const emailFieldValue:
      string | number | undefined =
      this.getStringOrNumberValue(
        item[
          configuration.emailInternalName
        ]
      );

    const contactNameFieldValue:
      string | number | undefined =
      this.getStringOrNumberValue(
        item[
          configuration.contactNameInternalName
        ]
      );

    const sustainabilityFieldValue:
      string | number | undefined =
      this.getStringOrNumberValue(
        item[
          configuration.overallPercentageInternalName
        ]
      );

    const weightingFieldValue:
      string | number | undefined =
      this.getStringOrNumberValue(
        item[
          configuration.weightingInternalName
        ]
      );

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
     * Displayed Sustainability percentage.
     *
     * Examples:
     * "80%" becomes 80.
     * "65%" becomes 65.
     *
     * A percent character is removed only for parsing.
     * The numeric scale is not changed.
     */
    const overallPercentage: number =
      SupplierQualificationUtility.parseSharePointValue(
        sustainabilityFieldValue
      );

    /**
     * Weighted qualification score.
     *
     * Examples:
     * "10" becomes 10.
     * "9.99" becomes 9.99.
     * "5" becomes 5.
     * "4.99" becomes 4.99.
     * "2" becomes 2.
     * "1.99" becomes 1.99.
     */
    const weightedScore: number =
      SupplierQualificationUtility.parseSharePointValue(
        weightingFieldValue
      );

    const qualificationResult: IQualificationResult =
      this.evaluateExactSharePointScore(
        weightedScore,
        configuration
      );

    return {
      key:
        `${configuration.tier}-` +
        `${item.Id.toString()}`,
      id: item.Id,
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
      overallPercentage,
      weightedScore,
      qualification:
        qualificationResult.qualification,
      riskRating:
        qualificationResult.riskRating,
      recommendation:
        qualificationResult.recommendation,
      created:
        item.Created,
      modified:
        item.Modified
    };
  }

  /**
   * Evaluates the exact SharePoint Weighting value
   * against the approved tier-specific thresholds.
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
    weightedScore: number,
    configuration: IListConfiguration
  ): IQualificationResult {
    if (
      !isFinite(
        weightedScore
      ) ||
      weightedScore < 0
    ) {
      return {
        qualification: 'Not Qualified',
        riskRating: 'High Risk',
        recommendation: 'Requires Review'
      };
    }

    if (
      weightedScore >=
      configuration.qualifiedWeightedMinimum
    ) {
      return {
        qualification: 'Qualified',
        riskRating: 'Low Risk',
        recommendation: 'Approve'
      };
    }

    if (
      weightedScore >=
      configuration.conditionalWeightedMinimum
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
      qualification: 'Not Qualified',
      riskRating: 'High Risk',
      recommendation: 'Requires Review'
    };
  }

  /**
   * Restricts a dynamically retrieved SharePoint value
   * to the types used by text and score fields.
   */
  private getStringOrNumberValue(
    value:
      string |
      number |
      boolean |
      null |
      undefined
  ): string | number | undefined {
    if (
      typeof value === 'string' ||
      typeof value === 'number'
    ) {
      return value;
    }

    return undefined;
  }

  /**
   * Converts a supported SharePoint field value
   * into displayable text.
   */
  private getTextValue(
    value:
      string |
      number |
      undefined
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

  /**
   * Removes duplicate field names from the REST
   * API select collection.
   */
  private getUniqueFieldNames(
    fieldNames: readonly string[]
  ): string[] {
    const uniqueFieldNames: string[] = [];

    fieldNames.forEach(
      (
        fieldName: string
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

  /**
   * Converts a date string into a numeric value
   * for submission sorting.
   */
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

  /**
   * Builds the SharePoint display form URL
   * for a submitted questionnaire item.
   */
  private buildItemUrl(
    listTitle: string,
    itemId: number
  ): string {
    return (
      `${this.webAbsoluteUrl}/Lists/` +
      `${encodeURIComponent(
        listTitle
      )}` +
      `/DispForm.aspx?ID=` +
      `${itemId.toString()}`
    );
  }

  /**
   * Escapes apostrophes used inside an OData
   * string value.
   */
  private escapeODataString(
    value: string
  ): string {
    return value.replace(
      /'/g,
      "''"
    );
  }
}