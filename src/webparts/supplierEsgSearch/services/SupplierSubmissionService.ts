/*import {
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
  IQualificationThresholds,
  SupplierQualificationUtility
} from '../utilities/SupplierQualificationUtility';

export class SupplierSubmissionService
  implements ISupplierSubmissionService {
  private readonly context: WebPartContext;
  private readonly webAbsoluteUrl: string;

  public constructor(context: WebPartContext) {
    this.context = context;

    this.webAbsoluteUrl =
      context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  public async getAllSubmissions(
    listConfigurations: readonly IListConfiguration[],
    thresholds: IQualificationThresholds
  ): Promise<ISupplierSubmission[]> {
    const listPromises:
      Array<Promise<ISupplierSubmission[]>> =
        listConfigurations.map(
          (
            configuration: IListConfiguration
          ): Promise<ISupplierSubmission[]> => {
            return this.getSubmissionsFromList(
              configuration,
              thresholds
            );
          }
        );

    const groupedResults: ISupplierSubmission[][] =
      await Promise.all(listPromises);

    let combinedResults: ISupplierSubmission[] = [];

    groupedResults.forEach(
      (group: ISupplierSubmission[]): void => {
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

  private async getSubmissionsFromList(
    configuration: IListConfiguration,
    thresholds: IQualificationThresholds
  ): Promise<ISupplierSubmission[]> {
    const encodedListTitle: string =
      this.escapeODataString(
        configuration.listTitle
      );

    const selectFields: string[] =
      this.getUniqueFieldNames([
        'Id',
        configuration.supplierNameInternalName,
        configuration.emailInternalName,
        configuration.contactNameInternalName,
        configuration.overallPercentageInternalName,
        'Created',
        'Modified'
      ]);

    const endpoint: string =
      `${this.webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle('${encodedListTitle}')/items` +
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
          configuration,
          thresholds
        );
      }
    );
  }

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
                'application/json;odata.metadata=none'
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

      const responseData:
        ISharePointItemsResponse =
          await response.json() as
            ISharePointItemsResponse;

      const pageItems:
        ISharePointSubmissionItem[] =
          responseData.value || [];

      allItems = allItems.concat(pageItems);

      nextEndpoint =
        responseData['@odata.nextLink'] ||
        responseData['odata.nextLink'];
    }

    return allItems;
  }

  private mapSubmission(
    item: ISharePointSubmissionItem,
    configuration: IListConfiguration,
    thresholds: IQualificationThresholds
  ): ISupplierSubmission {
    const supplierFieldValue:
      string | number | undefined =
        item[
          configuration.supplierNameInternalName
        ];

    const emailFieldValue:
      string | number | undefined =
        item[
          configuration.emailInternalName
        ];

    const contactNameFieldValue:
      string | number | undefined =
        item[
          configuration.contactNameInternalName
        ];

    const percentageFieldValue:
      string | number | undefined =
        item[
          configuration
            .overallPercentageInternalName
        ];

    const supplierName: string =
      this.getTextValue(supplierFieldValue);

    const email: string =
      this.getTextValue(emailFieldValue);

    const submittedByName: string =
      this.getTextValue(
        contactNameFieldValue
      );

    const overallPercentage: number =
      SupplierQualificationUtility
        .normalizePercentage(
          percentageFieldValue
        );

    const qualificationResult:
      IQualificationResult =
        SupplierQualificationUtility.evaluate(
          overallPercentage,
          thresholds
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

      overallPercentage,

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

  private getTextValue(
    value: string | number | undefined
  ): string {
    if (typeof value === 'string') {
      return value.trim();
    }

    if (typeof value === 'number') {
      return value.toString();
    }

    return '';
  }

  private getUniqueFieldNames(
    fieldNames: readonly string[]
  ): string[] {
    const uniqueFieldNames: string[] = [];

    fieldNames.forEach(
      (fieldName: string): void => {
        if (
          fieldName &&
          uniqueFieldNames.indexOf(fieldName) === -1
        ) {
          uniqueFieldNames.push(fieldName);
        }
      }
    );

    return uniqueFieldNames;
  }

  private getDateSortValue(
    dateValue: string
  ): number {
    const timeValue: number =
      new Date(dateValue).getTime();

    return isNaN(timeValue)
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
      `/DispForm.aspx?ID=${itemId.toString()}`
    );
  }

  private escapeODataString(
    value: string
  ): string {
    return value.replace(/'/g, "''");
  }
}*/

/*
*New Servuce to be used for the Supplier ESG Search Web Part. This service will be used to fetch the supplier submissions from the SharePoint lists and return the data to the web part. */
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
  IQualificationThresholds,
  SupplierQualificationUtility
} from '../utilities/SupplierQualificationUtility';

export class SupplierSubmissionService
  implements ISupplierSubmissionService {
  private readonly context: WebPartContext;
  private readonly webAbsoluteUrl: string;

  public constructor(context: WebPartContext) {
    this.context = context;

    this.webAbsoluteUrl =
      context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  /**
   * Retrieves submissions from all configured ESG questionnaire lists.
   */
  public async getAllSubmissions(
    listConfigurations: readonly IListConfiguration[],
    thresholds: IQualificationThresholds
  ): Promise<ISupplierSubmission[]> {
    const listPromises:
      Array<Promise<ISupplierSubmission[]>> =
      listConfigurations.map(
        (
          configuration: IListConfiguration
        ): Promise<ISupplierSubmission[]> => {
          return this.getSubmissionsFromList(
            configuration,
            thresholds
          );
        }
      );

    const groupedResults: ISupplierSubmission[][] =
      await Promise.all(listPromises);

    let combinedResults: ISupplierSubmission[] = [];

    groupedResults.forEach(
      (group: ISupplierSubmission[]): void => {
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
    configuration: IListConfiguration,
    thresholds: IQualificationThresholds
  ): Promise<ISupplierSubmission[]> {
    const encodedListTitle: string =
      this.escapeODataString(
        configuration.listTitle
      );

    /*
     * Each list can use different internal field names.
     *
     * Tier 1:
     * Email = field_3
     * Contact Name = field_4
     *
     * Tier 2 and Tier 3:
     * Email = Email
     * Contact Name = Name
     */
    const selectFields: string[] =
      this.getUniqueFieldNames([
        'Id',
        configuration.supplierNameInternalName,
        configuration.emailInternalName,
        configuration.contactNameInternalName,
        configuration.overallPercentageInternalName,
        'Created',
        'Modified'
      ]);

    const endpoint: string =
      `${this.webAbsoluteUrl}` +
      `/_api/web/lists/getbytitle('${encodedListTitle}')/items` +
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
        item: ISharePointSubmissionItem
      ): ISupplierSubmission => {
        return this.mapSubmission(
          item,
          configuration,
          thresholds
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
      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          nextEndpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept:
                'application/json;odata.metadata=none'
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

      const responseData:
        ISharePointItemsResponse =
        await response.json() as
        ISharePointItemsResponse;

      const pageItems:
        ISharePointSubmissionItem[] =
        responseData.value || [];

      allItems =
        allItems.concat(pageItems);

      nextEndpoint =
        responseData['@odata.nextLink'] ||
        responseData['odata.nextLink'];
    }

    return allItems;
  }

  /**
   * Converts one SharePoint item into the common supplier model.
   */
  private mapSubmission(
    item: ISharePointSubmissionItem,
    configuration: IListConfiguration,
    thresholds: IQualificationThresholds
  ): ISupplierSubmission {
    const supplierFieldValue:
      string | number | undefined =
      item[
      configuration.supplierNameInternalName
      ];

    const emailFieldValue:
      string | number | undefined =
      item[
      configuration.emailInternalName
      ];

    const contactNameFieldValue:
      string | number | undefined =
      item[
      configuration.contactNameInternalName
      ];

    const percentageFieldValue:
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

    /*
 * Preserve the exact SharePoint weighted value.
 *
 * Examples:
 * Tier 1 maximum = 10
 * Tier 2 maximum = 5
 * Tier 3 maximum = 2
 */
    const overallPercentage: number =
      this.parseSharePointCalculatedValue(
        percentageFieldValue
      );

    /*
     * Convert the weighted value to a 0–100 scale only
     * for qualification, risk, and recommendation.
     *
     * The displayed SharePoint value remains unchanged.
     */
    const qualificationPercentage: number =
      this.calculateQualificationPercentage(
        overallPercentage,
        configuration.maximumWeightedScore
      );

    const qualificationResult:
      IQualificationResult =
      SupplierQualificationUtility.evaluate(
        qualificationPercentage,
        thresholds
      );

    const submission:
      ISupplierSubmission = {
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

      overallPercentage,

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

    return submission;
  }

  private calculateQualificationPercentage(
    sharePointValue: number,
    maximumWeightedScore: number
  ): number {
    if (
      !isFinite(sharePointValue) ||
      sharePointValue < 0 ||
      !isFinite(maximumWeightedScore) ||
      maximumWeightedScore <= 0
    ) {
      return 0;
    }

    const relativePercentage: number =
      (
        sharePointValue /
        maximumWeightedScore
      ) * 100;

    return Math.min(
      Math.max(relativePercentage, 0),
      100
    );
  }

  private parseSharePointCalculatedValue(
    value: string | number | undefined
  ): number {
    if (typeof value === 'number') {
      return isFinite(value)
        ? value
        : 0;
    }

    if (typeof value !== 'string') {
      return 0;
    }

    const normalizedValue: string =
      value
        .replace('%', '')
        .replace(',', '.')
        .trim();

    if (!normalizedValue) {
      return 0;
    }

    const parsedValue: number =
      parseFloat(normalizedValue);

    if (
      isNaN(parsedValue) ||
      !isFinite(parsedValue)
    ) {
      return 0;
    }

    return parsedValue;
  }

  /**
   * Converts a dynamic SharePoint value to text.
   */
  private getTextValue(
    value: string | number | undefined
  ): string {
    if (typeof value === 'string') {
      return value.trim();
    }

    if (typeof value === 'number') {
      return value.toString();
    }

    return '';
  }

  /**
   * Removes duplicate internal field names before creating $select.
   */
  private getUniqueFieldNames(
    fieldNames: readonly string[]
  ): string[] {
    const uniqueFieldNames: string[] = [];

    fieldNames.forEach(
      (fieldName: string): void => {
        if (
          fieldName &&
          uniqueFieldNames.indexOf(fieldName) === -1
        ) {
          uniqueFieldNames.push(fieldName);
        }
      }
    );

    return uniqueFieldNames;
  }

  /**
   * Safely converts a SharePoint date into a sortable timestamp.
   */
  private getDateSortValue(
    dateValue: string
  ): number {
    const timeValue: number =
      new Date(dateValue).getTime();

    return isNaN(timeValue)
      ? 0
      : timeValue;
  }

  /**
   * Creates the SharePoint display-form URL.
   */
  private buildItemUrl(
    listTitle: string,
    itemId: number
  ): string {
    return (
      `${this.webAbsoluteUrl}/Lists/` +
      `${encodeURIComponent(listTitle)}` +
      `/DispForm.aspx?ID=${itemId.toString()}`
    );
  }

  /**
   * Escapes an apostrophe in the SharePoint list title.
   */
  private escapeODataString(
    value: string
  ): string {
    return value.replace(/'/g, "''");
  }
}