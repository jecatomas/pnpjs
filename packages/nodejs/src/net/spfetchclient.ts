declare var global: any;
import { HttpClientImpl, combine, isUrlAbsolute } from "@pnp/common";
import { NodeFetchClient } from "./nodefetchclient";
import { getAddInOnlyAccessToken } from "../sptokenutils";
import { SPOAuthEnv, AuthToken } from "../types";
import * as HttpsProxyAgent from "https-proxy-agent";

/**
 * Fetch client for use within nodejs, requires you register a client id and secret with app only permissions
 */
export class SPFetchClient  implements HttpClientImpl  {

    protected agent: HttpsProxyAgent;
    protected token: AuthToken | null = null;

    constructor(
        public siteUrl: string,
        protected _clientId: string,
        protected _clientSecret: string,
        public authEnv: SPOAuthEnv = SPOAuthEnv.SPO,
        protected _realm = "",
        protected _fetchClient: HttpClientImpl = new NodeFetchClient(),
        protected proxyUrl?: string) {

        global._spPageContextInfo = {
            webAbsoluteUrl: siteUrl,
        };
        this.agent = this.proxyUrl && new HttpsProxyAgent(this.proxyUrl);
    }

    public async fetch(url: string, options: any = {}): Promise<Response> {

        const realm = await this.getRealm();
        const authUrl = await this.getAuthUrl(realm);
        const token = await getAddInOnlyAccessToken(this.siteUrl, this._clientId, this._clientSecret, realm, authUrl, this.proxyUrl);

        options.headers.set("Authorization", `Bearer ${token.access_token}`);

        const uri = !isUrlAbsolute(url) ? combine(this.siteUrl, url) : url;

        return this._fetchClient.fetch(uri, options);
    }

    public getAuthHostUrl(env: SPOAuthEnv): string {
        switch (env) {
            case SPOAuthEnv.China:
                return "accounts.accesscontrol.chinacloudapi.cn";
            case SPOAuthEnv.Germany:
                return "login.microsoftonline.de";
            default:
                return "accounts.accesscontrol.windows.net";
        }
    }

    private async getRealm(): Promise<string> {

        if (this._realm.length > 0) {
            return Promise.resolve(this._realm);
        }

        const url = combine(this.siteUrl, "_vti_bin/client.svc");

        const r = await this._fetchClient.fetch(url, {
            "headers": {
                "Authorization": "Bearer ",
            },
            "method": "POST",
        });

        const data: string = r.headers.get("www-authenticate") || "";
        const index = data.indexOf("Bearer realm=\"");
        this._realm = data.substring(index + 14, index + 50);
        return this._realm;
    }

    private async getAuthUrl(realm: string): Promise<string> {

        const url = `https://${this.getAuthHostUrl(this.authEnv)}/metadata/json/1?realm=${realm}`;

        const r = await this._fetchClient.fetch(url, {
            agent: this.agent,
            method: "GET",
        });
        const json: { endpoints: { protocol: string, location: string }[] } = await r.json();

        const eps = json.endpoints.filter(ep => ep.protocol === "OAuth2");
        if (eps.length > 0) {
            return eps[0].location;
        }

        throw Error("Auth URL Endpoint could not be determined from data.");
    }
}
