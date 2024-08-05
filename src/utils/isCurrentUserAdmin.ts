import { spfi,SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import * as configs from 'GlobalSetting';

export default class isCurrentUserAdmin {
    private baseUrl = configs.Domain + configs.BaseUrl  + '/Documents%20partages/';
    private sp: ReturnType<typeof spfi>;

    constructor(private context: any) {
        this.sp = spfi().using(SPFx(this.context));
    }

    public async isAdmin(): Promise<boolean> {
        const adminResponse = await this.fetchJson(`${this.baseUrl}/admin.json`);
        const adminEmails = Object.values(adminResponse.adminEmails);
        const currentUser = await this.sp.web.currentUser();
        // Si l'utilisateur est admin par son email ou par son statut alors on retourne true
        return adminEmails.includes(this.context.pageContext.user.email) ||  currentUser.IsSiteAdmin;
    }

    private async fetchJson(url: string): Promise<any> {
        const response = await fetch(url, {
            headers: {
                'Accept': 'application/json;odata=verbose'
            }
        });
    
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
    
        return response.json();
    }
}
