import * as React from 'react';

import {
  Icon,
  IconButton
} from '@fluentui/react';

import {
  ISupplierNavigationProps
} from './ISupplierNavigationProps';

import styles from './SupplierNavigation.module.scss';

interface INavigationItem {
  key: string;
  title: string;
  iconName: string;
  url: string;
}

export function SupplierNavigation(
  props: ISupplierNavigationProps
): React.ReactElement<ISupplierNavigationProps> {
  const [
    isMobileMenuOpen,
    setIsMobileMenuOpen
  ] = React.useState<boolean>(false);

  const navigationItems: readonly INavigationItem[] = [
    {
      key: 'home',
      title: 'HOME',
      iconName: 'Home',
      url: props.homeUrl
    },
    {
      key: 'completedQuestionnaires',
      title: 'COMPLETED QUESTIONNARIES',
      iconName: 'PageList',
      url: props.completedQuestionnairesUrl
    },
    {
      key: 'supportingDocuments',
      title: 'SUPPORTING DOCUMENTS',
      iconName: 'Attach',
      url: props.supportingDocumentsUrl
    },
    {
      key: 'dashboard',
      title: 'DASHBOARD',
      iconName: 'BarChartVertical',
      url: props.dashboardUrl
    },
    {
      key: 'help',
      title: 'HELP',
      iconName: 'Help',
      url: props.helpUrl
    }
  ];

  const currentPath: string = normalizePath(
    window.location.pathname
  );

  const openNavigationItem = (
    item: INavigationItem
  ): void => {
    if (!item.url) {
      return;
    }

    setIsMobileMenuOpen(false);

    window.location.href = resolveUrl(item.url);
  };

  const isNavigationItemActive = (
    item: INavigationItem
  ): boolean => {
    const targetPath: string = getPathFromUrl(
      item.url
    );

    return currentPath === targetPath;
  };

  return (
    <nav
      className={styles.navigation}
      aria-label="Supplier Sustainability navigation"
    >
      <div className={styles.navigationInner}>
        <a
          href={resolveUrl(props.homeUrl || '#')}
          className={styles.brand}
          aria-label="Supplier Sustainability Surveys home"
          onClick={(
            event: React.MouseEvent<HTMLAnchorElement>
          ): void => {
            if (props.homeUrl) {
              event.preventDefault();
              openNavigationItem({
                key: 'home',
                title: 'HOME',
                iconName: 'Home',
                url: props.homeUrl
              });
            }
          }}
        >
        </a>

        <div
          className={styles.desktopNavigation}
          aria-label="Primary navigation"
        >
          {navigationItems.map(
            (
              item: INavigationItem
            ): React.ReactElement => {
              const isActive: boolean =
                isNavigationItemActive(item);

              return (
                <a
                  key={item.key}
                  href={resolveUrl(item.url)}
                  className={`${styles.navigationLink} ${
                    isActive ? styles.activeLink : ''
                  }`}
                  aria-current={
                    isActive ? 'page' : undefined
                  }
                  title={item.title}
                  onClick={(
                    event: React.MouseEvent<HTMLAnchorElement>
                  ): void => {
                    event.preventDefault();
                    openNavigationItem(item);
                  }}
                >
                  <span
                    className={styles.linkIcon}
                    aria-hidden="true"
                  >
                    <Icon
                      iconName={item.iconName}
                    />
                  </span>

                  <span
                    className={styles.linkText}
                  >
                    {item.title}
                  </span>
                </a>
              );
            }
          )}
        </div>

        <div className={styles.mobileMenuButton}>
          <IconButton
            iconProps={{
              iconName: isMobileMenuOpen
                ? 'Cancel'
                : 'GlobalNavButton'
            }}
            title={
              isMobileMenuOpen
                ? 'Close navigation menu'
                : 'Open navigation menu'
            }
            ariaLabel={
              isMobileMenuOpen
                ? 'Close navigation menu'
                : 'Open navigation menu'
            }
            aria-expanded={isMobileMenuOpen}
            onClick={(): void => {
              setIsMobileMenuOpen(
                (currentValue: boolean): boolean => {
                  return !currentValue;
                }
              );
            }}
          />
        </div>
      </div>

      {isMobileMenuOpen && (
        <div
          className={styles.mobileNavigation}
          aria-label="Mobile navigation"
        >
          {navigationItems.map(
            (
              item: INavigationItem
            ): React.ReactElement => {
              const isActive: boolean =
                isNavigationItemActive(item);

              return (
                <a
                  key={item.key}
                  href={resolveUrl(item.url)}
                  className={`${styles.mobileNavigationLink} ${
                    isActive ? styles.activeLink : ''
                  }`}
                  aria-current={
                    isActive ? 'page' : undefined
                  }
                  title={item.title}
                  onClick={(
                    event: React.MouseEvent<HTMLAnchorElement>
                  ): void => {
                    event.preventDefault();
                    openNavigationItem(item);
                  }}
                >
                  <span
                    className={styles.mobileLinkIcon}
                    aria-hidden="true"
                  >
                    <Icon
                      iconName={item.iconName}
                    />
                  </span>

                  <span>{item.title}</span>
                </a>
              );
            }
          )}
        </div>
      )}
    </nav>
  );
}

function resolveUrl(url: string): string {
  const normalizedUrl: string = (url || '').trim();

  if (!normalizedUrl) {
    return window.location.href;
  }

  const lowerUrl: string = normalizedUrl.toLowerCase();

  const isAbsoluteUrl: boolean =
    lowerUrl.indexOf('http://') === 0 ||
    lowerUrl.indexOf('https://') === 0;

  if (isAbsoluteUrl) {
    return normalizedUrl;
  }

  const relativeUrl: string =
    normalizedUrl.charAt(0) === '/'
      ? normalizedUrl
      : `/${normalizedUrl}`;

  return `${window.location.origin}${relativeUrl}`;
}

function getPathFromUrl(url: string): string {
  if (!url) {
    return '';
  }

  const anchor: HTMLAnchorElement =
    document.createElement('a');

  anchor.href = resolveUrl(url);

  return normalizePath(anchor.pathname);
}

function normalizePath(value: string): string {
  const normalizedValue: string = (value || '/')
    .toLowerCase()
    .replace(/\/+$/, '');

  return normalizedValue || '/';
}

export default SupplierNavigation;