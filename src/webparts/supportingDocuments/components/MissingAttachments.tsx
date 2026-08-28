import * as React from 'react';

import {
  DefaultButton,
  DetailsListLayoutMode,
  DetailsRow,
  Dropdown,
  IColumn,
  IDropdownOption,
  IDetailsRowProps,
  Icon,
  IconButton,
  Link,
  SelectionMode,
  ShimmeredDetailsList
} from '@fluentui/react';

import {
  IMissingAttachment
} from '../models/IMissingAttachment';

import styles from './MissingAttachments.module.scss';

export interface IMissingAttachmentsProps {
  items: readonly IMissingAttachment[];

  isLoading: boolean;

  onOpenSubmission(
    item: IMissingAttachment
  ): void;

  onViewAll?(): void;
}

const DEFAULT_PAGE_SIZE: number = 10;

const PAGE_SIZE_OPTIONS:
  readonly IDropdownOption[] = [
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

export function MissingAttachments(
  props: IMissingAttachmentsProps
): React.ReactElement<IMissingAttachmentsProps> {
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
      props.items
    ]
  );

  React.useEffect(
    (): void => {
      if (
        currentPage > totalPages
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
    IMissingAttachment[] =
      props.items
        .slice(
          startIndex,
          endIndex
        );

  const formatDate = (
    value: string
  ): string => {
    if (!value) {
      return '-';
    }

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
        return styles.defaultTier;
    }
  };

  const openSubmission = (
    item?: IMissingAttachment
  ): void => {
    if (
      !item ||
      props.isLoading
    ) {
      return;
    }

    props.onOpenSubmission(
      item
    );
  };

  const columns: IColumn[] = [
    {
      key: 'supplierName',
      name: 'Supplier',
      fieldName: 'supplierName',
      minWidth: 145,
      maxWidth: 220,
      isResizable: true,

      onRender: (
        item: IMissingAttachment
      ): React.ReactNode => {
        return (
          <div
            className={styles.supplierCell}
            title={item.supplierName}
          >
            <span
              className={styles.warningIcon}
              aria-hidden="true"
            >
              <Icon iconName="Warning" />
            </span>

            <div className={styles.supplierText}>
              <Link
                className={styles.supplierName}
                title={
                  `Open submission for ` +
                  item.supplierName
                }
                onClick={(
                  event:
                    React.MouseEvent<
                      HTMLElement
                    >
                ): void => {
                  event.preventDefault();
                  event.stopPropagation();

                  openSubmission(
                    item
                  );
                }}
              >
                {item.supplierName}
              </Link>

              <span
                className={
                  styles.supplierListLabel
                }
              >
                Supplier List Item
              </span>
            </div>
          </div>
        );
      }
    },
    {
      key: 'tier',
      name: 'Tier',
      fieldName: 'tier',
      minWidth: 56,
      maxWidth: 72,
      isResizable: true,

      onRender: (
        item: IMissingAttachment
      ): React.ReactNode => {
        return (
          <span
            className={
              `${styles.tierBadge} ` +
              getTierClassName(
                item.tier
              )
            }
          >
            {item.tier}
          </span>
        );
      }
    },
    {
      key: 'submissionDate',
      name: 'Submitted',
      fieldName: 'submissionDate',
      minWidth: 82,
      maxWidth: 110,
      isResizable: true,

      onRender: (
        item: IMissingAttachment
      ): React.ReactNode => {
        return (
          <span
            className={
              styles.submissionDate
            }
          >
            {formatDate(
              item.submissionDate
            )}
          </span>
        );
      }
    },
    {
      key: 'status',
      name: 'Status',
      minWidth: 68,
      maxWidth: 82,
      isResizable: false,

      onRender: (
        _item: IMissingAttachment
      ): React.ReactNode => {
        return (
          <span
            className={
              styles.missingBadge
            }
          >
            Missing
          </span>
        );
      }
    }
  ];

  const renderRow = (
    rowProps?: IDetailsRowProps
  ): JSX.Element | null => {
    if (!rowProps) {
      return null;
    }

    const item:
      IMissingAttachment =
        rowProps.item as
          IMissingAttachment;

    return (
      <div
        className={styles.clickableRow}
        role="button"
        tabIndex={
          props.isLoading
            ? -1
            : 0
        }
        aria-label={
          `Open submission for ` +
          item.supplierName
        }
        aria-disabled={
          props.isLoading
        }
        onClick={(): void => {
          openSubmission(
            item
          );
        }}
        onKeyDown={(
          event:
            React.KeyboardEvent<
              HTMLDivElement
            >
        ): void => {
          if (
            event.key !== 'Enter' &&
            event.key !== ' '
          ) {
            return;
          }

          event.preventDefault();

          openSubmission(
            item
          );
        }}
      >
        <DetailsRow {...rowProps} />
      </div>
    );
  };

  const goToPage = (
    pageNumber: number
  ): void => {
    const safePageNumber: number =
      Math.min(
        Math.max(
          pageNumber,
          1
        ),
        totalPages
      );

    setCurrentPage(
      safePageNumber
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
            typeof page === 'string'
          ) {
            return (
              <span
                key={page}
                className={styles.ellipsis}
                aria-hidden="true"
              >
                ...
              </span>
            );
          }

          const isCurrentPage:
            boolean =
              page === currentPage;

          return (
            <DefaultButton
              key={page}
              className={
                isCurrentPage
                  ? styles.activePageButton
                  : styles.pageButton
              }
              text={page.toString()}
              ariaLabel={
                isCurrentPage
                  ? (
                    `Current page, ` +
                    `page ${page}`
                  )
                  : `Go to page ${page}`
              }
              disabled={
                props.isLoading
              }
              onClick={(): void => {
                goToPage(
                  page
                );
              }}
            />
          );
        }
      );
    };

  const openAllItems = (): void => {
    if (
      props.isLoading ||
      totalItems === 0 ||
      !props.onViewAll
    ) {
      return;
    }

    props.onViewAll();
  };

  return (
    <section
      className={styles.panel}
      aria-labelledby=
        "missing-attachments-title"
    >
      <header className={styles.header}>
        <div>
          <h2 id="missing-attachments-title">
            Missing Attachments
          </h2>

          <span className={styles.headerSubtitle}>
            Questionnaire submitted without
            supporting documents
          </span>
        </div>

        <span
          className={styles.countBadge}
          aria-label={
            `${totalItems} missing attachment` +
            `${totalItems === 1 ? '' : 's'}`
          }
        >
          {totalItems}
        </span>
      </header>

      {props.isLoading ||
      visibleItems.length > 0 ? (
        <div className={styles.tableScroll}>
          <ShimmeredDetailsList
            items={visibleItems}
            columns={columns}
            setKey={
              `missing-attachments-` +
              `${currentPage}-` +
              `${pageSize}`
            }
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
            onRenderRow={renderRow}
            ariaLabelForGrid={
              'Suppliers with missing attachments'
            }
            ariaLabelForShimmer={
              'Checking questionnaire submissions ' +
              'for missing attachments'
            }
          />
        </div>
      ) : (
        <div
          className={styles.message}
          role="status"
        >
          <span
            className={styles.completeIcon}
            aria-hidden="true"
          >
            <Icon iconName="Completed" />
          </span>

          <strong>
            No missing attachments
          </strong>

          <p>
            Every submitted questionnaire
            currently has at least one
            supporting document.
          </p>
        </div>
      )}

      <footer className={styles.footer}>
        <div
          className={styles.paginationButtons}
          aria-label=
            "Missing attachment pagination"
        >
          <IconButton
            iconProps={{
              iconName: 'ChevronLeft'
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
              iconName: 'ChevronRight'
            }}
            title="Next page"
            ariaLabel="Next page"
            disabled={
              props.isLoading ||
              currentPage === totalPages ||
              totalItems === 0
            }
            onClick={(): void => {
              goToPage(
                currentPage + 1
              );
            }}
          />
        </div>

        <div className={styles.footerActions}>
          <span className={styles.footerCount}>
            {totalItems === 0
              ? 'No items'
              : (
                `Showing ` +
                `${startIndex + 1}` +
                `–${endIndex} of ` +
                `${totalItems}`
              )}
          </span>

          <div className={styles.pageSizeSelector}>
            <span>
              Items per page:
            </span>

            <Dropdown
              ariaLabel="Items per page"
              options={
                PAGE_SIZE_OPTIONS as
                  IDropdownOption[]
              }
              selectedKey={pageSize}
              disabled={props.isLoading}
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

          {props.onViewAll && (
            <Link
              disabled={
                props.isLoading ||
                totalItems === 0
              }
              onClick={
                openAllItems
              }
            >
              View all
            </Link>
          )}
        </div>
      </footer>
    </section>
  );
}

export default MissingAttachments;