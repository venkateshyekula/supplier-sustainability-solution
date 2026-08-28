import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {
  IDocumentLibraryConfiguration
} from '../models/IDocumentLibraryConfiguration';

import {
  ISupplierDocument
} from '../models/ISupplierDocument';

import {
  ISupplierDocumentService
} from './ISupplierDocumentService';

interface IUser {
  Title?: string;
  EMail?: string;
}

interface IFile {
  Name?: string;
  Length?: string;
  ServerRelativeUrl?: string;
  TimeCreated?: string;
  TimeLastModified?: string;
}

interface IItem {
  Id: number;
  FileDirRef?: string;
  Created?: string;
  Modified?: string;
  Author?: IUser;
  Editor?: IUser;
  File?: IFile;
}

interface IResponse {
  value?: IItem[];

  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

interface IMeta {
  RootFolder?: {
    ServerRelativeUrl?: string;
  };
}

export class SupplierDocumentService
  implements ISupplierDocumentService {

  private readonly context: WebPartContext;
  private readonly webUrl: string;

  public constructor(
    context: WebPartContext
  ) {
    this.context = context;

    this.webUrl =
      context.pageContext.web.absoluteUrl
        .replace(
          /\/$/,
          ''
        );
  }

  /**
   * Retrieves documents from all configured Tier document
   * libraries and combines them into one result collection.
   */
  public async getAllDocuments(
    configurations:
      readonly IDocumentLibraryConfiguration[]
  ): Promise<ISupplierDocument[]> {
    const documentGroups:
      ISupplierDocument[][] =
        await Promise.all(
          configurations.map(
            (
              configuration:
                IDocumentLibraryConfiguration
            ): Promise<ISupplierDocument[]> => {
              return this.getDocumentsFromLibrary(
                configuration
              );
            }
          )
        );

    let allDocuments:
      ISupplierDocument[] = [];

    /*
     * Use a block-bodied callback so the assignment is not
     * implicitly returned. This resolves ESLint no-return-assign.
     */
    documentGroups.forEach(
      (
        documentGroup:
          ISupplierDocument[]
      ): void => {
        allDocuments =
          allDocuments.concat(
            documentGroup
          );
      }
    );

    allDocuments.sort(
      (
        first:
          ISupplierDocument,

        second:
          ISupplierDocument
      ): number => {
        return (
          this.getTimeValue(
            second.created
          ) -
          this.getTimeValue(
            first.created
          )
        );
      }
    );

    return allDocuments;
  }

  /**
   * Resolves the root folder URL for each configured document
   * library.
   *
   * The returned object uses Tier 1, Tier 2, and Tier 3 as keys.
   */
  public async getLibraryUrls(
    configurations:
      readonly IDocumentLibraryConfiguration[]
  ): Promise<Record<string, string>> {
    const libraryUrls:
      Record<string, string> = {};

    await Promise.all(
      configurations.map(
        async (
          configuration:
            IDocumentLibraryConfiguration
        ): Promise<void> => {
          const libraryRootUrl:
            string =
              await this.getLibraryRootUrl(
                configuration.libraryTitle
              );

          libraryUrls[
            configuration.tier
          ] = libraryRootUrl;
        }
      )
    );

    return libraryUrls;
  }

  /**
   * Retrieves all files from one configured SharePoint document
   * library.
   *
   * Folder items are excluded through FSObjType eq 0.
   */
  private async getDocumentsFromLibrary(
    configuration:
      IDocumentLibraryConfiguration
  ): Promise<ISupplierDocument[]> {
    const libraryRootUrl:
      string =
        await this.getLibraryRootUrl(
          configuration.libraryTitle
        );

    const escapedLibraryTitle:
      string =
        this.escapeODataString(
          configuration.libraryTitle
        );

    const endpoint:
      string =
        `${this.webUrl}` +
        `/_api/web/lists/getbytitle(` +
        `'${escapedLibraryTitle}')/items` +
        `?$select=` +
        [
          'Id',
          'FileDirRef',
          'Created',
          'Modified',
          'Author/Title',
          'Author/EMail',
          'Editor/Title',
          'Editor/EMail',
          'File/Name',
          'File/Length',
          'File/ServerRelativeUrl',
          'File/TimeCreated',
          'File/TimeLastModified'
        ].join(',') +
        `&$expand=Author,Editor,File` +
        `&$filter=FSObjType eq 0` +
        `&$orderby=Created desc` +
        `&$top=5000`;

    const items:
      IItem[] =
        await this.getAllPages(
          endpoint,
          configuration.libraryTitle
        );

    return items
      .filter(
        (
          item:
            IItem
        ): boolean => {
          return Boolean(
            item.File &&
            item.File.ServerRelativeUrl
          );
        }
      )
      .map(
        (
          item:
            IItem
        ): ISupplierDocument => {
          const fileName:
            string =
              item.File &&
              item.File.Name
                ? item.File.Name
                : 'Unnamed file';

          const folderServerRelativeUrl:
            string =
              item.FileDirRef ||
              libraryRootUrl;

          const fileServerRelativeUrl:
            string =
              item.File &&
              item.File.ServerRelativeUrl
                ? item.File.ServerRelativeUrl
                : '';

          const created:
            string =
              item.File &&
              item.File.TimeCreated
                ? item.File.TimeCreated
                : item.Created || '';

          const modified:
            string =
              item.File &&
              item.File.TimeLastModified
                ? item.File.TimeLastModified
                : item.Modified || '';

          const documentStatus:
  ISupplierDocument ['status'] = 'Submitted';

return {
  key:
    `${configuration.tier}-` +
    `${item.Id.toString()}`,

  itemId:
    item.Id,

  supplierName:
    this.getSupplierNameFromFolder(
      libraryRootUrl,
      folderServerRelativeUrl
    ),

  tier:
    configuration.tier,

  libraryTitle:
    configuration.libraryTitle,

  fileName,

  fileExtension:
    this.getFileExtension(
      fileName
    ),

  fileSize:
    this.parseFileSize(
      item.File
        ? item.File.Length
        : undefined
    ),

  created,

  createdBy:
    this.getUserDisplayValue(
      item.Author
    ),

  modified,

  modifiedBy:
    this.getUserDisplayValue(
      item.Editor
    ),

  folderServerRelativeUrl,

  fileServerRelativeUrl,

  status:
    documentStatus
};
        }
      );
  }

  /**
   * Resolves the server-relative root URL of a configured
   * SharePoint document library.
   */
  private async getLibraryRootUrl(
    libraryTitle: string
  ): Promise<string> {
    const escapedLibraryTitle:
      string =
        this.escapeODataString(
          libraryTitle
        );

    const endpoint:
      string =
        `${this.webUrl}` +
        `/_api/web/lists/getbytitle(` +
        `'${escapedLibraryTitle}')` +
        `?$select=RootFolder/ServerRelativeUrl` +
        `&$expand=RootFolder`;

    const response:
      SPHttpClientResponse =
        await this.context
          .spHttpClient
          .get(
            endpoint,
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
        `Unable to resolve document library ` +
        `"${libraryTitle}". ` +
        `HTTP status: ${response.status} ` +
        `${response.statusText}. ` +
        `Response: ${responseText}`
      );
    }

    const metadata:
      IMeta =
        await response.json() as IMeta;

    const libraryRootUrl:
      string =
        metadata.RootFolder &&
        metadata.RootFolder
          .ServerRelativeUrl
          ? metadata.RootFolder
              .ServerRelativeUrl
          : '';

    if (!libraryRootUrl) {
      throw new Error(
        `SharePoint did not return a root folder URL ` +
        `for document library "${libraryTitle}".`
      );
    }

    return libraryRootUrl.replace(
      /\/$/,
      ''
    );
  }

  /**
   * Retrieves all SharePoint REST pages.
   */
  private async getAllPages(
    initialEndpoint: string,
    libraryTitle: string
  ): Promise<IItem[]> {
    let allItems:
      IItem[] = [];

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
          `Unable to retrieve documents from ` +
          `"${libraryTitle}". ` +
          `HTTP status: ${response.status} ` +
          `${response.statusText}. ` +
          `Response: ${responseText}`
        );
      }

      const responseData:
        IResponse =
          await response.json() as
          IResponse;

      const pageItems:
        IItem[] =
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
   * Gets the supplier name from the first folder immediately
   * below the document-library root.
   *
   * Example:
   *
   * Library root:
   * /sites/SupplierSustainability/TestSupply
   *
   * File directory:
   * /sites/SupplierSustainability/TestSupply/ABC Engineering
   *
   * Result:
   * ABC Engineering
   */
  private getSupplierNameFromFolder(
    libraryRootUrl: string,
    folderServerRelativeUrl: string
  ): string {
    const normalizedLibraryRoot:
      string =
        this.safeDecodeURIComponent(
          libraryRootUrl.replace(
            /\/$/,
            ''
          )
        );

    const normalizedFolderUrl:
      string =
        this.safeDecodeURIComponent(
          folderServerRelativeUrl
            .replace(
              /\/$/,
              ''
            )
        );

    if (
      normalizedFolderUrl ===
      normalizedLibraryRoot
    ) {
      return 'Library Root';
    }

    const expectedPrefix:
      string =
        `${normalizedLibraryRoot}/`;

    const relativeFolderPath:
      string =
        normalizedFolderUrl.indexOf(
          expectedPrefix
        ) === 0
          ? normalizedFolderUrl.substring(
              expectedPrefix.length
            )
          : normalizedFolderUrl;

    const folderParts:
      string[] =
        relativeFolderPath
          .split('/')
          .map(
            (
              folderPart:
                string
            ): string => {
              return folderPart.trim();
            }
          )
          .filter(
            (
              folderPart:
                string
            ): boolean => {
              return Boolean(
                folderPart
              );
            }
          );

    return folderParts.length > 0
      ? folderParts[0]
      : 'Unknown Supplier';
  }

  private getFileExtension(
    fileName: string
  ): string {
    const extensionSeparatorIndex:
      number =
        fileName.lastIndexOf('.');

    if (
      extensionSeparatorIndex < 0 ||
      extensionSeparatorIndex ===
        fileName.length - 1
    ) {
      return 'FILE';
    }

    return fileName
      .substring(
        extensionSeparatorIndex + 1
      )
      .toUpperCase();
  }

  private parseFileSize(
    value?: string
  ): number {
    if (!value) {
      return 0;
    }

    const parsedValue:
      number =
        parseInt(
          value,
          10
        );

    if (
      isNaN(parsedValue) ||
      !isFinite(parsedValue) ||
      parsedValue < 0
    ) {
      return 0;
    }

    return parsedValue;
  }

  private getUserDisplayValue(
    user?: IUser
  ): string {
    if (!user) {
      return 'Unknown';
    }

    return (
      user.Title ||
      user.EMail ||
      'Unknown'
    );
  }

  private getTimeValue(
    dateValue: string
  ): number {
    const timeValue:
      number =
        new Date(
          dateValue
        ).getTime();

    return isNaN(timeValue)
      ? 0
      : timeValue;
  }

  private escapeODataString(
    value: string
  ): string {
    return value.replace(
      /'/g,
      "''"
    );
  }

  /**
   * Prevents malformed encoded folder names from causing
   * decodeURIComponent to throw an exception.
   */
  private safeDecodeURIComponent(
    value: string
  ): string {
    try {
      return decodeURIComponent(
        value
      );
    } catch {
      return value;
    }
  }
}