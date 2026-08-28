import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {
  spfi,
  SPFI,
  SPFx
} from '@pnp/sp';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';

let spInstance:
  SPFI | undefined;

let currentWebUrl:
  string | undefined;

export function getSP(
  context: WebPartContext
): SPFI {
  const webAbsoluteUrl: string =
    context.pageContext.web.absoluteUrl;

  if (
    !spInstance ||
    currentWebUrl !== webAbsoluteUrl
  ) {
    spInstance =
      spfi()
        .using(
          SPFx(context)
        );

    currentWebUrl =
      webAbsoluteUrl;
  }

  return spInstance;
}