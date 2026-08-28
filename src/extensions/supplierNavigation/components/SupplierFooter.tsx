import * as React from 'react';

import styles from './SupplierFooter.module.scss';

export interface ISupplierFooterProps {
  organizationName?: string;
  confidentialityText?: string;
}

export function SupplierFooter(
  props: ISupplierFooterProps
): React.ReactElement<ISupplierFooterProps> {
  const currentYear: number =
    new Date().getFullYear();

  const organizationName: string =
    props.organizationName &&
    props.organizationName.trim()
      ? props.organizationName.trim()
      : 'Cassava Technologies';

  const confidentialityText: string =
    props.confidentialityText &&
    props.confidentialityText.trim()
      ? props.confidentialityText.trim()
      : (
        'All information is confidential and must be used ' +
        'in accordance with our privacy and security policies.'
      );

  return (
    <footer
      className={styles.footer}
      role="contentinfo"
      aria-label="Supplier Sustainability footer"
    >
      <div className={styles.footerContent}>
        <p className={styles.confidentiality}>
          {confidentialityText}
        </p>

        <p className={styles.copyright}>
          © {currentYear} {organizationName}. All rights reserved.
        </p>
      </div>
    </footer>
  );
}

export default SupplierFooter;