import * as pnp from 'sp-pnp-js';
import NodeFetchClient from 'node-pnp-js';

export class PnPConfig {
    public static configure(siteUrl: string): void {
        pnp.setup({
            sp: {
                fetchClientFactory: () => {
                    return new NodeFetchClient({
                        ondemand: true
                    });
                },
                baseUrl: siteUrl
            }
        });
    }
}
