import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IQuestionnaireQuickLinksProps {
  context: WebPartContext;

  tier1QuestionnaireListTitle: string;
  tier2QuestionnaireListTitle: string;
  tier3QuestionnaireListTitle: string;

  tier1DocumentLibraryTitle: string;
  tier2DocumentLibraryTitle: string;
  tier3DocumentLibraryTitle: string;
}
