import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/attachments";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";
import "@pnp/sp/sputilities";
import "@pnp/sp/profiles"
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";

var _sp: any = null;

const getSP = (context: WebPartContext): SPFI => {
  _sp = spfi().using(SPFx(context));
  return _sp;
};

export default getSP;
