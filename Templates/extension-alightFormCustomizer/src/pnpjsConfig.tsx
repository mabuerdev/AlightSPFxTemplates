/* eslint-disable no-var */
// import pnp and pnp logging system
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import { ISPFXContext, SPFx as spSPFx } from "@pnp/sp/behaviors/spfx";
import { spfi, SPFI } from "@pnp/sp/fi";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/comments/clientside-page";
import "@pnp/sp/comments/item";
import "@pnp/sp/content-types";
import "@pnp/sp/site-users/web"
import "@pnp/sp/regional-settings/web"
import "@pnp/sp/profiles";
import "@pnp/sp/search";
import "@pnp/sp/fields";
import "@pnp/sp/site-users";
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "@pnp/sp/files";
import "@pnp/sp/site-groups/web";

import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/graph/sites"
import "@pnp/graph/lists";
import "@pnp/graph/list-item";

import { IWeb, Web } from "@pnp/sp/webs";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";

let _sp: SPFI;
let _graph: GraphFI;

export const getSP = (context?: FormCustomizerContext): SPFI => {
  if (context) {
    console.log("Initializing PNP SP context");
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(spSPFx(context as ISPFXContext));
  }

  return _sp;
};

export const getGraph = (context?: FormCustomizerContext): GraphFI => {
  if (context) {
    console.log("Initializing Graph context");
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _graph = graphfi().using(graphSPFx(context as ISPFXContext));
  }
  
  return _graph;
};

export const getContextForUrl = (sp: SPFI, foreignUrl: string): IWeb | undefined => {
    if (_sp !== null) {
        return Web([sp.web, foreignUrl]);
    }

    return undefined;
}

export const getContextForWebId = async (sp:SPFI, webId: string): Promise<IWeb | undefined> => {
    if (_sp !== null) {
        return (await sp.site.openWebById(webId)).web;
    }

    return undefined;
}