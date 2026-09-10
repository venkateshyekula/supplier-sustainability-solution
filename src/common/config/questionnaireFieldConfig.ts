export type QuestionnaireTier =
  | 'Tier 1'
  | 'Tier 2'
  | 'Tier 3';

export interface IQuestionnaireFieldMapping {
  overallPercentageInternalName:
    string;

  weightingInternalName:
    string;
}

interface IQuestionnaireEnvironmentMapping {
  tier1:
    IQuestionnaireFieldMapping;

  tier2:
    IQuestionnaireFieldMapping;

  tier3:
    IQuestionnaireFieldMapping;
}

const PREPROD_SITE_URL:
  string =
    'https://liquidtelecommunications.sharepoint.com/sites/ltsadev/sc';

const PRODUCTION_SITE_URL:
  string =
    'https://liquidtelecommunications.sharepoint.com/sites/sss';

const PREPROD_FIELDS:
  IQuestionnaireEnvironmentMapping = {
    tier1: {
      overallPercentageInternalName:
        'OverallQuestionsPercentage',

      weightingInternalName:
        'OverallQuestionsPercentage'
    },

    tier2: {
      overallPercentageInternalName:
        'OverallQuestionsPercentage',

      weightingInternalName:
        'OverallQuestionsPercentage'
    },

    tier3: {
      overallPercentageInternalName:
        'OverallQuestionsPercentage',

      weightingInternalName:
        'OverallQuestionsPercentage'
    }
  };

const PRODUCTION_FIELDS:
  IQuestionnaireEnvironmentMapping = {
    tier1: {
      overallPercentageInternalName:
        'Tier_x0020_1_x0020_Sustainabilit',

      weightingInternalName:
        'Tier_x0020_1_x0020_Weighting'
    },

    tier2: {
      overallPercentageInternalName:
        'Tier_x0020_2_x0020_Sustainabilit',

      weightingInternalName:
        'Tier_x0020_2_x0020_Weighting'
    },

    tier3: {
      overallPercentageInternalName:
        'Tier_x0020_3_x0020_Sustainabilit',

      weightingInternalName:
        'Tier_x0020_3_x0020_Weighting'
    }
  };

export function getQuestionnaireFields(
  webAbsoluteUrl:
    string,

  tier:
    QuestionnaireTier
): IQuestionnaireFieldMapping {
  const normalizedWebUrl:
    string =
      normalizeWebUrl(
        webAbsoluteUrl
      );

  const normalizedPreProdUrl:
    string =
      normalizeWebUrl(
        PREPROD_SITE_URL
      );

  const normalizedProductionUrl:
    string =
      normalizeWebUrl(
        PRODUCTION_SITE_URL
      );

  let environmentMapping:
    IQuestionnaireEnvironmentMapping;

  if (
    normalizedWebUrl ===
    normalizedProductionUrl
  ) {
    environmentMapping =
      PRODUCTION_FIELDS;
  } else if (
    normalizedWebUrl ===
    normalizedPreProdUrl
  ) {
    environmentMapping =
      PREPROD_FIELDS;
  } else {
    throw new Error(
      'Questionnaire fields are not configured ' +
      `for SharePoint web: ${webAbsoluteUrl}`
    );
  }

  switch (tier) {
    case 'Tier 1':
      return environmentMapping.tier1;

    case 'Tier 2':
      return environmentMapping.tier2;

    case 'Tier 3':
      return environmentMapping.tier3;

    default:
      throw new Error(
        `Unsupported questionnaire tier: ${tier}`
      );
  }
}

function normalizeWebUrl(
  value:
    string
): string {
  return (
    value ||
    ''
  )
    .trim()
    .replace(
      /\/+$/,
      ''
    )
    .toLowerCase();
}