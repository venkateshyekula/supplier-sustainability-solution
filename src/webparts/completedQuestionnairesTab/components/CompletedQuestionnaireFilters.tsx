import * as React from 'react';
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  SearchBox
} from '@fluentui/react';
import { QualificationStatus } from '../../supplierEsgSearch/models/IQualificationResult';
import { ISupplierSearchFilters } from '../../supplierEsgSearch/models/ISupplierSearchFilters';
import { SupplierTier } from '../../supplierEsgSearch/models/IListConfiguration';
import { ICompletedQuestionnaireFiltersProps } from './ICompletedQuestionnaireFiltersProps';
import styles from './CompletedQuestionnaireFilters.module.scss';

const TIER_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'All Tiers' },
  { key: 'Tier 1', text: 'Tier 1' },
  { key: 'Tier 2', text: 'Tier 2' },
  { key: 'Tier 3', text: 'Tier 3' }
];

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: '', text: 'All ESG Statuses' },
  { key: 'Qualified', text: 'Qualified' },
  { key: 'Conditionally Qualified', text: 'Conditionally Qualified' },
  { key: 'Not Qualified', text: 'Not Qualified' }
];

export function CompletedQuestionnaireFilters(
  props: ICompletedQuestionnaireFiltersProps
): React.ReactElement<ICompletedQuestionnaireFiltersProps> {
  const hasCriteria: boolean = Boolean(
    props.filters.searchText.trim() ||
    props.filters.tier ||
    props.filters.qualification ||
    props.filters.startDate ||
    props.filters.endDate
  );

  const update = (changes: Partial<ISupplierSearchFilters>): void => {
    props.onFiltersChange({ ...props.filters, ...changes });
  };

  return (
    <section className={styles.filters} aria-labelledby="completed-search-title">
      <h2 id="completed-search-title">Search &amp; Filter Submissions</h2>
      <div className={styles.grid}>
        <SearchBox
          className={styles.search}
          placeholder="Search by supplier name, contact or email..."
          value={props.filters.searchText}
          disabled={props.isLoading}
          onChange={(_event?: React.ChangeEvent<HTMLInputElement>, value?: string): void => {
            update({ searchText: value || '' });
          }}
          onSearch={(): void => props.onSearch()}
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
        <Dropdown
          options={STATUS_OPTIONS}
          selectedKey={props.filters.qualification || ''}
          disabled={props.isLoading}
          onChange={(_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
            update({
              qualification: option && option.key
                ? option.key as QualificationStatus
                : undefined
            });
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
        <PrimaryButton
          text="Search"
          iconProps={{ iconName: 'Search' }}
          disabled={props.isLoading || !hasCriteria}
          onClick={props.onSearch}
        />
        <DefaultButton
          text="Clear"
          disabled={props.isLoading || !hasCriteria}
          onClick={props.onClear}
        />
      </div>
    </section>
  );
}

export default CompletedQuestionnaireFilters;