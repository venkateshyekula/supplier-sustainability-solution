import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {
  ISiteSearchResult
} from '../models/ISiteSearchResult';

interface ISearchCell {
  Key: string;

  Value:
    string |    
    undefined;

  ValueType?:
    string;
}

interface ISearchCellsContainer {
  results?:
    ISearchCell[];
}

interface ISearchRow {
  Cells?:
    ISearchCell[] |
    ISearchCellsContainer;
}

interface ISearchRowsContainer {
  results?:
    ISearchRow[];
}

interface ISearchTable {
  Rows?:
    ISearchRow[] |
    ISearchRowsContainer;
}

interface ISearchRelevantResults {
  Table?:
    ISearchTable;

  RowCount?:
    number;

  TotalRows?:
    number;

  TotalRowsIncludingDuplicates?:
    number;
}

interface ISearchPrimaryQueryResult {
  RelevantResults?:
    ISearchRelevantResults;
}

interface ISearchQueryResult {
  PrimaryQueryResult?:
    ISearchPrimaryQueryResult;
}

interface ISearchResponse {
  d?: {
    query?:
      ISearchQueryResult;
  };

  PrimaryQueryResult?:
    ISearchPrimaryQueryResult;
}

export class SiteSearchService {
  private readonly context:
    WebPartContext;

  private readonly webAbsoluteUrl:
    string;

  private readonly siteAbsoluteUrl:
    string;

  public constructor(
    context:
      WebPartContext
  ) {
    this.context =
      context;

    /*
     * REST requests are sent through the current web.
     */
    this.webAbsoluteUrl =
      context
        .pageContext
        .web
        .absoluteUrl
        .replace(
          /\/+$/,
          ''
        );

    /*
     * Search results are restricted to the entire
     * current site collection, not only the current web.
     */
    this.siteAbsoluteUrl =
      (
        context
          .pageContext
          .site
          .absoluteUrl ||
        context
          .pageContext
          .web
          .absoluteUrl
      )
        .replace(
          /\/+$/,
          ''
        );
  }

  /**
   * Searches indexed content in the current SharePoint
   * site collection.
   *
   * SharePoint Search automatically applies security
   * trimming for the logged-in user.
   */
  public async searchSite(
    searchText:
      string,

    rowLimit:
      number = 30
  ): Promise<ISiteSearchResult[]> {
    const normalizedSearchText:
      string =
        searchText.trim();

    if (
      normalizedSearchText.length < 2
    ) {
      return [];
    }

    const safeRowLimit:
      number =
        this.getSafeRowLimit(
          rowLimit
        );

    const queryText:
      string =
        this.buildQueryText(
          normalizedSearchText
        );

    const selectProperties:
      string = [
        'Title',
        'Path',
        'OriginalPath',
        'FileType',
        'FileExtension',
        'ContentClass',
        'ContentType',
        'HitHighlightedSummary',
        'ServerRedirectedURL',
        'SiteName',
        'WebId',
        'ListId',
        'ListItemID'
      ].join(',');

    const endpoint:
      string =
        `${this.webAbsoluteUrl}` +
        '/_api/search/query' +
        `?querytext='${encodeURIComponent(
          queryText
        )}'` +
        `&selectproperties='${encodeURIComponent(
          selectProperties
        )}'` +
        `&rowlimit=${safeRowLimit.toString()}` +
        '&startrow=0' +
        '&trimduplicates=false' +
        '&enablequeryrules=true' +
        '&processbestbets=true';

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
        'Unable to search the SharePoint site. ' +
        `HTTP ${response.status.toString()} ` +
        `${response.statusText}. ` +
        responseText
      );
    }

    const responseData:
      ISearchResponse =
        await response.json() as
          ISearchResponse;

    const primaryQueryResult:
      ISearchPrimaryQueryResult |
      undefined =
        responseData
          .PrimaryQueryResult ||
        responseData
          .d
          ?.query
          ?.PrimaryQueryResult;

    const relevantResults:
      ISearchRelevantResults |
      undefined =
        primaryQueryResult
          ?.RelevantResults;

    const rows:
      ISearchRow[] =
        this.getSearchRows(
          relevantResults
            ?.Table
            ?.Rows
        );

    return rows
      .map(
        (
          row:
            ISearchRow,

          index:
            number
        ):
          ISiteSearchResult |
          undefined => {
          return this.mapSearchResult(
            row,
            index
          );
        }
      )
      .filter(
        (
          result:
            ISiteSearchResult |
            undefined
        ): result is ISiteSearchResult => {
          return Boolean(
            result?.path
          );
        }
      );
  }

  /**
   * Creates the SharePoint KQL query.
   *
   * The wildcard provides prefix matching and Path
   * limits results to the current site collection.
   */
  private buildQueryText(
    searchText:
      string
  ): string {
    const escapedSearchText:
      string =
        this.escapeKqlValue(
          searchText
        );

    const escapedSiteUrl:
      string =
        this.escapeKqlValue(
          this.siteAbsoluteUrl
        );

    return (
      `${escapedSearchText}* ` +
      `AND Path:"${escapedSiteUrl}"`
    );
  }

  /**
   * Supports both minimal-metadata and verbose SharePoint
   * Search response formats.
   */
  private getSearchRows(
    rows:
      ISearchRow[] |
      ISearchRowsContainer |
      undefined
  ): ISearchRow[] {
    if (
      Array.isArray(
        rows
      )
    ) {
      return rows;
    }

    return rows?.results || [];
  }

  /**
   * Supports both minimal-metadata and verbose cell
   * response formats.
   */
  private getSearchCells(
    cells:
      ISearchCell[] |
      ISearchCellsContainer |
      undefined
  ): ISearchCell[] {
    if (
      Array.isArray(
        cells
      )
    ) {
      return cells;
    }

    return cells?.results || [];
  }

  private mapSearchResult(
    row:
      ISearchRow,

    index:
      number
  ):
    ISiteSearchResult |
    undefined {
    const cells:
      ISearchCell[] =
        this.getSearchCells(
          row.Cells
        );

    const values:
      Record<string, string> = {};

    cells.forEach(
      (
        cell:
          ISearchCell
      ): void => {
        values[cell.Key] =
          cell.Value || '';
      }
    );

    const path:
      string =
        (
          values.ServerRedirectedURL ||
          values.OriginalPath ||
          values.Path ||
          ''
        )
          .trim();

    if (!path) {
      return undefined;
    }

    const title:
      string =
        this.cleanSearchText(
          values.Title
        ) ||
        this.getFallbackTitle(
          path
        );

    const fileType:
      string =
        (
          values.FileType ||
          values.FileExtension ||
          ''
        )
          .trim()
          .toLowerCase();

    const contentClass:
      string =
        (
          values.ContentClass ||
          ''
        )
          .trim();

    const summary:
      string =
        this.cleanSearchSummary(
          values
            .HitHighlightedSummary
        );

    return {
      key:
        `${path}-${index.toString()}`,

      title,

      path,

      summary,

      fileType,

      contentClass,

      iconName:
        this.getIconName(
          fileType,
          contentClass,
          path
        )
    };
  }

  private cleanSearchSummary(
    value:
      string | undefined
  ): string {
    if (!value) {
      return '';
    }

    const cleanedValue:
      string =
        value
          .replace(
            /<ddd\/>/gi,
            '...'
          )
          .replace(
            /<c\d+>/gi,
            ''
          )
          .replace(
            /<\/c\d+>/gi,
            ''
          );

    return this.cleanSearchText(
      cleanedValue
    );
  }

  private cleanSearchText(
    value:
      string | undefined
  ): string {
    if (!value) {
      return '';
    }

    const element:
      HTMLDivElement =
        document.createElement(
          'div'
        );

    element.innerHTML =
      value;

    return (
      element.textContent ||
      element.innerText ||
      ''
    )
      .replace(
        /\s+/g,
        ' '
      )
      .trim();
  }

  private getFallbackTitle(
    path:
      string
  ): string {
    try {
      const url:
        URL =
          new URL(
            path
          );

      const pathParts:
        string[] =
          url
            .pathname
            .split('/')
            .filter(
              (
                part:
                  string
              ): boolean => {
                return Boolean(
                  part
                );
              }
            );

      const finalPart:
        string =
          pathParts[
            pathParts.length - 1
          ] ||
          'SharePoint result';

      return decodeURIComponent(
        finalPart
      );
    } catch {
      return 'SharePoint result';
    }
  }

  private getIconName(
    fileType:
      string,

    contentClass:
      string,

    path:
      string
  ): string {
    const normalizedPath:
      string =
        path
          .trim()
          .toLowerCase();

    const normalizedContentClass:
      string =
        contentClass
          .trim()
          .toLowerCase();

    if (
      normalizedContentClass.indexOf(
        'sts_site'
      ) >= 0 ||
      normalizedContentClass.indexOf(
        'sts_web'
      ) >= 0
    ) {
      return 'SharepointLogo';
    }

    if (
      normalizedPath.indexOf(
        '/sitepages/'
      ) >= 0 ||
      fileType === 'aspx'
    ) {
      return 'Page';
    }

    switch (fileType) {
      case 'doc':
      case 'docx':
        return 'WordDocument';

      case 'xls':
      case 'xlsx':
      case 'csv':
        return 'ExcelDocument';

      case 'ppt':
      case 'pptx':
        return 'PowerPointDocument';

      case 'pdf':
        return 'PDF';

      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
      case 'svg':
        return 'Photo2';

      case 'zip':
      case 'rar':
      case '7z':
        return 'ZipFolder';

      case 'txt':
        return 'TextDocument';

      default:
        return 'Document';
    }
  }

  private getSafeRowLimit(
    rowLimit:
      number
  ): number {
    if (
      !Number.isFinite(
        rowLimit
      ) ||
      rowLimit < 1
    ) {
      return 30;
    }

    return Math.min(
      Math.floor(
        rowLimit
      ),
      100
    );
  }

  private escapeKqlValue(
    value:
      string
  ): string {
    return value
      .replace(
        /\\/g,
        '\\\\'
      )
      .replace(
        /"/g,
        '\\"'
      )
      .replace(
        /'/g,
        "''"
      );
  }
}