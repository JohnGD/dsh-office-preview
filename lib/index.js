//#region lib/types/index.js
/**
 * Pure host half.
 *
 * The document pane, its file-address routing and the byte transport all belong
 * to `@deepseek-ai/dsh-client-ui-sidebar-documentpreview`; this package only
 * adds two renderer registrations on the browser side. The host plugin exists so
 * the Cordis loader carries a client entry whose `dsh.client` metadata the
 * module system can compose — it contributes nothing to the host tree.
 */

/** Host plugin body: no host-side service, tool, or route. */
function apply() {}

export { apply };
//#endregion
