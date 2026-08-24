/*import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import '@pnp/sp/folders';

let _sp: SPFI | undefined = undefined;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== undefined) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp!;
};*/

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
import '@pnp/sp/folders';

let spInstance: SPFI | undefined;

export const getSP = (
  context: WebPartContext
): SPFI => {
  if (!spInstance) {
    spInstance = spfi().using(
      SPFx(context)
    );
  }

  return spInstance;
};