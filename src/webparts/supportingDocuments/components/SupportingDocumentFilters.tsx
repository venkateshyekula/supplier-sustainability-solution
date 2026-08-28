import * as React from 'react';
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  SearchBox,
  TextField
} from '@fluentui/react';
import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';
import { ISupplierDocumentFilters } from '../models/ISupplierDocumentFilters';
import styles from './SupportingDocumentFilters.module.scss';

export interface ISupportingDocumentFiltersProps {
  filters: ISupplierDocumentFilters;
  isLoading: boolean;
  onFiltersChange(filters: ISupplierDocumentFilters): void;
  onSearch(): void;
  onClear(): void;
}

const TIER_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'All Tiers' },
  { key: 'Tier 1', text: 'Tier 1' },
  { key: 'Tier 2', text: 'Tier 2' },
  { key: 'Tier 3', text: 'Tier 3' }
];

export function SupportingDocumentFilters(
  props: ISupportingDocumentFiltersProps
): React.ReactElement<ISupportingDocumentFiltersProps> {
  const hasCriteria: boolean = Boolean(
    props.filters.searchText.trim() ||
    props.filters.tier ||
    props.filters.uploadedBy?.trim() ||
    props.filters.startDate ||
    props.filters.endDate
  );

  const update = (changes: Partial<ISupplierDocumentFilters>): void => {
    props.onFiltersChange({ ...props.filters, ...changes });
  };

  return (
    <section className={styles.filters} aria-labelledby="supporting-document-filters-title">
      <h2 id="supporting-document-filters-title">Search &amp; Filter Documents</h2>
      <div className={styles.grid}>
        <SearchBox
          placeholder="Search supplier or file name..."
          value={props.filters.searchText}
          disabled={props.isLoading}
          onChange={(_event?: React.ChangeEvent<HTMLInputElement>, value?: string): void => {
            update({ searchText: value || '' });
          }}
          onSearch={props.onSearch}
          onClear={(): void => update({ searchText: '' })}
        />
        <Dropdown
          options={TIER_OPTIONS}
          selectedKey={props.filters.tier || ''}
          disabled={props.isLoading}
          onChange={(_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
            update({ tier: option && option.key ? option.key as SupplierTier : undefined });
          }}
        />
        <TextField
          placeholder="Uploaded by"
          value={props.filters.uploadedBy || ''}
          disabled={props.isLoading}
          onChange={(_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
            update({ uploadedBy: value || undefined });
          }}
        />
        <DatePicker
          placeholder="From date"
          value={props.filters.startDate}
          disabled={props.isLoading}
          onSelectDate={(date: Date | null | undefined): void => update({ startDate: date || undefined })}
        />
        <DatePicker
          placeholder="To date"
          value={props.filters.endDate}
          disabled={props.isLoading}
          onSelectDate={(date: Date | null | undefined): void => update({ endDate: date || undefined })}
        />
        <PrimaryButton text="Search" disabled={props.isLoading || !hasCriteria} onClick={props.onSearch} />
        <DefaultButton text="Clear" disabled={props.isLoading || !hasCriteria} onClick={props.onClear} />
      </div>
    </section>
  );
}

export default SupportingDocumentFilters;