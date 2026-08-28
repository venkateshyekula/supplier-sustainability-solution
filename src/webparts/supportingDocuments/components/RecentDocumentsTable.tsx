import * as React from 'react';

import {
  DefaultButton,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  IDropdownOption,
  Icon,
  IconButton,
  Link,
  SelectionMode,
  ShimmeredDetailsList
} from '@fluentui/react';

import {
  ISupplierDocument
} from '../models/ISupplierDocument';

import styles from './RecentDocumentsTable.module.scss';

export interface IRecentDocumentsTableProps {
  items: readonly ISupplierDocument[];
  isLoading: boolean;

  onOpenDocument(
    item: ISupplierDocument
  ): void;
}

const DEFAULT_PAGE_SIZE: number = 10;

const PAGE_SIZE_OPTIONS:
  IDropdownOption[] = [
    {
      key: 5,
      text: '5'
    },
    {
      key: 10,
      text: '10'
    },
    {
      key: 15,
      text: '15'
    },
    {
      key: 20,
      text: '20'
    }
  ];

export function RecentDocumentsTable(
  props: IRecentDocumentsTableProps
): React.ReactElement<IRecentDocumentsTableProps> {
  const [
    currentPage,
    setCurrentPage
  ] = React.useState<number>(1);

  const [
    pageSize,
    setPageSize
  ] = React.useState<number>(
    DEFAULT_PAGE_SIZE
  );

  const totalItems: number =
    props.items.length;

  const totalPages: number =
    Math.max(
      1,
      Math.ceil(
        totalItems /
        pageSize
      )
    );

  React.useEffect(
    (): void => {
      setCurrentPage(1);
    },
    [
      props.items,
      pageSize
    ]
  );

  React.useEffect(
    (): void => {
      if (
        currentPage >
        totalPages
      ) {
        setCurrentPage(
          totalPages
        );
      }
    },
    [
      currentPage,
      totalPages
    ]
  );

  const startIndex: number =
    (
      currentPage -
      1
    ) *
    pageSize;

  const endIndex: number =
    Math.min(
      startIndex +
      pageSize,
      totalItems
    );

  const visibleItems:
    readonly ISupplierDocument[] =
      props.items.slice(
        startIndex,
        endIndex
      );

  const formatDate = (
    value: string
  ): string => {
    const date: Date =
      new Date(value);

    if (
      isNaN(
        date.getTime()
      )
    ) {
      return '-';
    }

    return date.toLocaleDateString(
      'en-GB',
      {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      }
    );
  };

  const getTierClassName = (
    tier: string
  ): string => {
    switch (tier) {
      case 'Tier 1':
        return styles.tier1;

      case 'Tier 2':
        return styles.tier2;

      case 'Tier 3':
        return styles.tier3;

      default:
        return '';
    }
  };

  const getFileIconName = (
    extension: string
  ): string => {
    switch (
      extension.toLowerCase()
    ) {
      case 'pdf':
        return 'PDF';

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

      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
        return 'Photo2';

      case 'zip':
      case 'rar':
      case '7z':
        return 'ZipFolder';

      default:
        return 'TextDocument';
    }
  };

  const columns: IColumn[] = [
    {
      key: 'file',
      name: 'File Name',
      fieldName: 'fileName',
      minWidth: 190,
      maxWidth: 290,
      isResizable: true,

      onRender: (
        item: ISupplierDocument
      ): React.ReactNode => {
        return (
          <Link
            className={
              styles.documentLink
            }
            title={item.fileName}
            onClick={(): void => {
              props.onOpenDocument(
                item
              );
            }}
          >
            <Icon
              className={
                styles.fileIcon
              }
              iconName={
                getFileIconName(
                  item.fileExtension
                )
              }
              aria-hidden="true"
            />

            <span>
              {item.fileName}
            </span>
          </Link>
        );
      }
    },
    {
      key: 'supplier',
      name: 'Supplier',
      fieldName: 'supplierName',
      minWidth: 150,
      maxWidth: 220,
      isResizable: true
    },
    {
      key: 'tier',
      name: 'Tier',
      fieldName: 'tier',
      minWidth: 65,
      maxWidth: 80,
      isResizable: true,

      onRender: (
        item: ISupplierDocument
      ): React.ReactNode => {
        return (
          <span
            className={
              `${styles.tierBadge} ` +
              `${getTierClassName(
                item.tier
              )}`
            }
          >
            {item.tier}
          </span>
        );
      }
    },
    {
      key: 'type',
      name: 'Type',
      fieldName: 'fileExtension',
      minWidth: 55,
      maxWidth: 75,
      isResizable: true
    },
    {
      key: 'uploadedBy',
      name: 'Uploaded By',
      fieldName: 'createdBy',
      minWidth: 120,
      maxWidth: 180,
      isResizable: true
    },
    {
      key: 'uploadedDate',
      name: 'Uploaded Date',
      fieldName: 'created',
      minWidth: 105,
      maxWidth: 135,
      isResizable: true,

      onRender: (
        item: ISupplierDocument
      ): React.ReactNode => {
        return formatDate(
          item.created
        );
      }
    },
    {
      key: 'size',
      name: 'Size',
      fieldName: 'fileSize',
      minWidth: 65,
      maxWidth: 90,
      isResizable: true,

      onRender: (
        item: ISupplierDocument
      ): React.ReactNode => {
        return formatBytes(
          item.fileSize
        );
      }
    },
    {
      key: 'actions',
      name: '',
      minWidth: 36,
      maxWidth: 36,
      isResizable: false,

      onRender: (
        item: ISupplierDocument
      ): React.ReactNode => {
        return (
          <IconButton
            iconProps={{
              iconName:
                'MoreVertical'
            }}
            title="Document actions"
            ariaLabel={
              `Actions for ` +
              `${item.fileName}`
            }
            menuProps={{
              items: [
                {
                  key: 'open',
                  text: 'Open document',
                  iconProps: {
                    iconName:
                      'OpenInNewWindow'
                  },
                  onClick: (): void => {
                    props.onOpenDocument(
                      item
                    );
                  }
                }
              ]
            }}
          />
        );
      }
    }
  ];

  const goToPage = (
    pageNumber: number
  ): void => {
    const safePage: number =
      Math.min(
        Math.max(
          pageNumber,
          1
        ),
        totalPages
      );

    setCurrentPage(
      safePage
    );
  };

  const renderPageButtons =
    (): React.ReactNode[] => {
      const pageNumbers:
        Array<number | string> = [];

      if (
        totalPages <= 7
      ) {
        for (
          let pageNumber: number = 1;
          pageNumber <= totalPages;
          pageNumber += 1
        ) {
          pageNumbers.push(
            pageNumber
          );
        }
      } else {
        pageNumbers.push(1);

        if (
          currentPage > 4
        ) {
          pageNumbers.push(
            'start-ellipsis'
          );
        }

        const firstMiddlePage:
          number =
            Math.max(
              2,
              currentPage - 1
            );

        const lastMiddlePage:
          number =
            Math.min(
              totalPages - 1,
              currentPage + 1
            );

        for (
          let pageNumber:
            number =
              firstMiddlePage;
          pageNumber <=
            lastMiddlePage;
          pageNumber += 1
        ) {
          pageNumbers.push(
            pageNumber
          );
        }

        if (
          currentPage <
          totalPages - 3
        ) {
          pageNumbers.push(
            'end-ellipsis'
          );
        }

        pageNumbers.push(
          totalPages
        );
      }

      return pageNumbers.map(
        (
          page:
            number | string
        ): React.ReactNode => {
          if (
            typeof page ===
            'string'
          ) {
            return (
              <span
                key={page}
                className={
                  styles.ellipsis
                }
                aria-hidden="true"
              >
                ...
              </span>
            );
          }

          return (
            <DefaultButton
              key={page}
              className={
                page ===
                currentPage
                  ? styles.activePageButton
                  : styles.pageButton
              }
              text={
                page.toString()
              }
              ariaLabel={
                `Go to page ${page}`
              }
              aria-current={
                page === currentPage
                  ? 'page'
                  : undefined
              }
              disabled={
                props.isLoading
              }
              onClick={(): void => {
                goToPage(page);
              }}
            />
          );
        }
      );
    };

  return (
    <section
      className={styles.panel}
      aria-labelledby=
        "recent-documents-title"
    >
      <header className={styles.header}>
        <div>
          <h2 id="recent-documents-title">
            Document Library
          </h2>

          <span>
            {totalItems === 0
              ? 'No documents'
              : (
                `Showing ` +
                `${startIndex + 1}` +
                `–${endIndex} of ` +
                `${totalItems} documents`
              )}
          </span>
        </div>
      </header>

      <div className={styles.scroll}>
        <ShimmeredDetailsList
          items={
            visibleItems as
              ISupplierDocument[]
          }
          columns={columns}
          setKey="supporting-documents"
          enableShimmer={
            props.isLoading
          }
          selectionMode={
            SelectionMode.none
          }
          layoutMode={
            DetailsListLayoutMode
              .justified
          }
          compact={true}
          ariaLabelForGrid=
            "Supporting documents"
          ariaLabelForShimmer=
            "Loading supporting documents"
        />
      </div>

      {!props.isLoading &&
        totalItems === 0 && (
          <div
            className={styles.empty}
          >
            No matching documents found.
          </div>
        )}

      <footer
        className={styles.pagination}
        aria-label=
          "Document table pagination"
      >
        <div
          className={
            styles.paginationButtons
          }
        >
          <IconButton
            iconProps={{
              iconName:
                'ChevronLeft'
            }}
            title="Previous page"
            ariaLabel="Previous page"
            disabled={
              props.isLoading ||
              currentPage === 1 ||
              totalItems === 0
            }
            onClick={(): void => {
              goToPage(
                currentPage - 1
              );
            }}
          />

          {renderPageButtons()}

          <IconButton
            iconProps={{
              iconName:
                'ChevronRight'
            }}
            title="Next page"
            ariaLabel="Next page"
            disabled={
              props.isLoading ||
              currentPage ===
                totalPages ||
              totalItems === 0
            }
            onClick={(): void => {
              goToPage(
                currentPage + 1
              );
            }}
          />
        </div>

        <div
          className={
            styles.pageSizeSelector
          }
        >
          <span>
            Items per page:
          </span>

          <Dropdown
            ariaLabel="Items per page"
            options={
              PAGE_SIZE_OPTIONS
            }
            selectedKey={
              pageSize
            }
            disabled={
              props.isLoading
            }
            onChange={(
              _event:
                React.FormEvent<
                  HTMLDivElement
                >,
              option?:
                IDropdownOption
            ): void => {
              if (!option) {
                return;
              }

              setPageSize(
                Number(
                  option.key
                )
              );

              setCurrentPage(1);
            }}
            styles={{
              root: {
                width: 72
              }
            }}
          />
        </div>
      </footer>
    </section>
  );
}

function formatBytes(
  value: number
): string {
  if (
    !isFinite(value) ||
    value <= 0
  ) {
    return '0 B';
  }

  const units: string[] = [
    'B',
    'KB',
    'MB',
    'GB'
  ];

  const index: number =
    Math.min(
      Math.floor(
        Math.log(value) /
        Math.log(1024)
      ),
      units.length - 1
    );

  return (
    `${
      (
        value /
        Math.pow(
          1024,
          index
        )
      ).toFixed(
        index === 0
          ? 0
          : 1
      )
    } ${units[index]}`
  );
}

export default RecentDocumentsTable;