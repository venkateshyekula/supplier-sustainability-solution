import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

let _sp: SPFI | undefined = undefined;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== undefined) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp!;
};