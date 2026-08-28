import * as React from "react";

import {
  DefaultButton,
  DetailsListLayoutMode,
  DetailsRow,
  Dropdown,
  IColumn,
  IDetailsRowProps,
  IDropdownOption,
  Icon,
  IconButton,
  Link,
  SelectionMode,
  ShimmeredDetailsList,
} from "@fluentui/react";

import { ISupplierFolderSummary } from "../models/ISupplierFolderSummary";

import styles from "./SupplierFolderTable.module.scss";

export interface ISupplierFolderTableProps {
  items: readonly ISupplierFolderSummary[];

  isLoading: boolean;

  onOpenFolder(item: ISupplierFolderSummary): void;
}

const DEFAULT_PAGE_SIZE: number = 10;

const PAGE_SIZE_OPTIONS: IDropdownOption[] = [
  {
    key: 5,
    text: "5",
  },
  {
    key: 10,
    text: "10",
  },
  {
    key: 15,
    text: "15",
  },
  {
    key: 20,
    text: "20",
  },
];

export function SupplierFolderTable(
  props: ISupplierFolderTableProps,
): React.ReactElement<ISupplierFolderTableProps> {
  const [currentPage, setCurrentPage] = React.useState<number>(1);

  const [pageSize, setPageSize] = React.useState<number>(DEFAULT_PAGE_SIZE);

  const totalItems: number = props.items.length;

  const totalPages: number = Math.max(1, Math.ceil(totalItems / pageSize));

  React.useEffect((): void => {
    setCurrentPage(1);
  }, [props.items, pageSize]);

  React.useEffect((): void => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages);
    }
  }, [currentPage, totalPages]);

  const startIndex: number = (currentPage - 1) * pageSize;

  const endIndex: number = Math.min(startIndex + pageSize, totalItems);

  const visibleItems: readonly ISupplierFolderSummary[] = props.items.slice(
    startIndex,
    endIndex,
  );

  const formatDate = (value?: string): string => {
    if (!value) {
      return "-";
    }

    const date: Date = new Date(value);

    if (isNaN(date.getTime())) {
      return "-";
    }

    return date.toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  };

  const columns: IColumn[] = [
    {
      key: "supplier",
      name: "Supplier Name",
      fieldName: "supplierName",
      minWidth: 170,
      maxWidth: 260,
      isResizable: true,

      onRender: (item: ISupplierFolderSummary): React.ReactNode => {
        return (
          <div className={styles.supplierCell}>
            <span className={styles.supplierIcon} aria-hidden="true">
              <Icon iconName="People" />
            </span>

            <div className={styles.supplierDetails}>
              <Link
                className={styles.supplierLink}
                title={item.supplierName}
                onClick={(): void => {
                  props.onOpenFolder(item);
                }}
              >
                {item.supplierName}
              </Link>

              <span className={styles.supplierFolderLabel}>
                Supplier folder
              </span>
            </div>
          </div>
        );
      },
    },
    {
      key: "tier",
      name: "Tier",
      fieldName: "tier",
      minWidth: 65,
      maxWidth: 80,
      isResizable: true,
    },
    {
      key: "count",
      name: "Documents",
      fieldName: "documentCount",
      minWidth: 75,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "size",
      name: "Total Size",
      fieldName: "totalFileSize",
      minWidth: 85,
      maxWidth: 110,
      isResizable: true,

      onRender: (item: ISupplierFolderSummary): React.ReactNode => {
        return formatBytes(item.totalFileSize);
      },
    },
    {
      key: "latest",
      name: "Latest Upload",
      fieldName: "latestUploadDate",
      minWidth: 105,
      maxWidth: 135,
      isResizable: true,

      onRender: (item: ISupplierFolderSummary): React.ReactNode => {
        return formatDate(item.latestUploadDate);
      },
    },
    {
      key: "modifiedBy",
      name: "Last Modified By",
      fieldName: "latestModifiedBy",
      minWidth: 125,
      maxWidth: 180,
      isResizable: true,
    },
  ];

  const renderDetailsRow = (
    rowProps?: IDetailsRowProps,
  ): JSX.Element | null => {
    if (!rowProps) {
      return null;
    }

    return (
      <DetailsRow
        {...rowProps}
        className={styles.detailsRow}
        styles={{
          root: {
            minHeight: 54,
            borderBottom: "1px solid #edebe9",
          },

          fields: {
            alignItems: "center",
          },

          cell: {
            display: "flex",
            alignItems: "center",
          },
        }}
      />
    );
  };

  const goToPage = (pageNumber: number): void => {
    setCurrentPage(Math.min(Math.max(pageNumber, 1), totalPages));
  };

  const renderPageButtons = (): React.ReactNode[] => {
    const pages: Array<number | string> = [];

    if (totalPages <= 7) {
      for (let page: number = 1; page <= totalPages; page += 1) {
        pages.push(page);
      }
    } else {
      pages.push(1);

      if (currentPage > 4) {
        pages.push("start-ellipsis");
      }

      const start: number = Math.max(2, currentPage - 1);

      const end: number = Math.min(totalPages - 1, currentPage + 1);

      for (let page: number = start; page <= end; page += 1) {
        pages.push(page);
      }

      if (currentPage < totalPages - 3) {
        pages.push("end-ellipsis");
      }

      pages.push(totalPages);
    }

    return pages.map((page: number | string): React.ReactNode => {
      if (typeof page === "string") {
        return (
          <span key={page} className={styles.ellipsis} aria-hidden="true">
            ...
          </span>
        );
      }

      return (
        <DefaultButton
          key={page}
          className={
            page === currentPage ? styles.activePageButton : styles.pageButton
          }
          text={page.toString()}
          aria-current={page === currentPage ? "page" : undefined}
          disabled={props.isLoading}
          onClick={(): void => {
            goToPage(page);
          }}
        />
      );
    });
  };

  return (
    <section className={styles.panel} aria-labelledby="supplier-folders-title">
      <header className={styles.header}>
        <h2 id="supplier-folders-title">Supplier Folders</h2>

        <span className={styles.resultSummary}>
          {props.isLoading
            ? "Loading folders..."
            : totalItems === 0
              ? "No folders"
              : `Showing ` +
                `${startIndex + 1}` +
                `–${endIndex} of ` +
                `${totalItems} folders`}
        </span>
      </header>

      <div className={styles.scroll}>
        <ShimmeredDetailsList
          items={visibleItems as ISupplierFolderSummary[]}
          columns={columns}
          setKey="supplier-folders"
          enableShimmer={props.isLoading}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          compact={true}
          ariaLabelForGrid="Supplier document folders"
          ariaLabelForShimmer={"Loading supplier document folders"}
          onRenderRow={renderDetailsRow}
        />
      </div>

      {!props.isLoading && totalItems === 0 && (
        <div className={styles.empty}>No supplier folders found.</div>
      )}

      <footer
        className={styles.pagination}
        aria-label="Supplier folder pagination"
      >
        <div className={styles.paginationButtons}>
          <IconButton
            iconProps={{
              iconName: "ChevronLeft",
            }}
            ariaLabel="Previous page"
            disabled={props.isLoading || currentPage === 1 || totalItems === 0}
            onClick={(): void => {
              goToPage(currentPage - 1);
            }}
          />

          {renderPageButtons()}

          <IconButton
            iconProps={{
              iconName: "ChevronRight",
            }}
            ariaLabel="Next page"
            disabled={
              props.isLoading || currentPage === totalPages || totalItems === 0
            }
            onClick={(): void => {
              goToPage(currentPage + 1);
            }}
          />
        </div>

        <div className={styles.pageSizeSelector}>
          <span>Items per page:</span>

          <Dropdown
            options={PAGE_SIZE_OPTIONS}
            selectedKey={pageSize}
            disabled={props.isLoading}
            onChange={(
              _event: React.FormEvent<HTMLDivElement>,
              option?: IDropdownOption,
            ): void => {
              if (!option) {
                return;
              }

              setPageSize(Number(option.key));

              setCurrentPage(1);
            }}
            styles={{
              root: {
                width: 72,
              },
            }}
          />
        </div>
      </footer>
    </section>
  );
}

function formatBytes(value: number): string {
  if (!isFinite(value) || value <= 0) {
    return "0 B";
  }

  const units: string[] = ["B", "KB", "MB", "GB"];

  const index: number = Math.min(
    Math.floor(Math.log(value) / Math.log(1024)),
    units.length - 1,
  );

  return `${(value / Math.pow(1024, index)).toFixed(
    index === 0 ? 0 : 1,
  )} ${units[index]}`;
}

export default SupplierFolderTable;
