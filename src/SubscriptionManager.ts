import * as pnp from 'sp-pnp-js';
import * as chalk from 'chalk';
import { List } from 'sp-pnp-js';

export class SubscriptionManager {

    constructor(private args: any) { }

    public async viewSubscriptions(verbose = true): Promise<any> {
        let listId = this.args['list'];
        let id = this.args['id'];

        if (!listId) {
            throw new Error(`Missing required parameter '--list'`);
        }

        let list = this.getList(listId);
        let data: any;
        if (id) {
            data = await list.subscriptions.getById(id).get();
        } else {
            data = await list.subscriptions.get();
        }
        if (verbose) {
            console.log('Subscription data received:');
            console.log(data);
        }

        return data;
    }

    public async addSubscription(): Promise<any> {
        let listId = this.args['list'];
        let url = this.args['url'];
        let expired = this.args['exp'];

        if (!listId) {
            throw new Error(`Missing required parameter '--list'`);
        }

        if (!url) {
            throw new Error(`Missing required parameter '--url'`);
        }

        if (!expired) {
            expired = this.getDefaultExpiration();
        }

        let list = this.getList(listId);
        let data = await list.subscriptions.add(url, expired);
        console.log('Subscription added:');
        console.log(data);

        return data;

    }

    public async deleteSubscription(id?: string): Promise<any> {
        let listId = this.args['list'];
        let subscriptionId = this.args['id'] || id;

        if (!listId) {
            throw new Error(`Missing required parameter '--list'`);
        }

        if (!subscriptionId) {
            throw new Error(`Missing required parameter '--id'`);
        }
        let list = this.getList(listId);
        await list.subscriptions.getById(subscriptionId).delete();
        console.log(`Subscription ${subscriptionId} deleted`);
    }

    public async deleteAllSubscriptions(): Promise<any> {

        let subscriptions = await this.viewSubscriptions(false);
        let deletePromises: any[] = [];
        subscriptions.forEach((subscription: any) => {
            deletePromises.push(this.deleteSubscription(subscription.id));
        });

        return Promise.all(deletePromises);
    }

    public async updateSubscription(): Promise<any> {

        let listId = this.args['list'];
        let expired = this.args['exp'];
        let subscriptionId = this.args['id'];

        if (!listId) {
            throw new Error(`Missing required parameter '--list'`);
        }

        if (!subscriptionId) {
            throw new Error(`Missing required parameter '--id'`);
        }

        if (!expired) {
            expired = this.getDefaultExpiration();
        }
        let list = this.getList(listId);
        await list.subscriptions.getById(subscriptionId).update(expired);

        console.log(`Subscription ${subscriptionId} updated`);
    }

    private getDefaultExpiration(): string {
        let now = new Date();
        now.setMonth(now.getMonth() + 6);
        return now.toISOString();
    }

    private getList(id: string): List {
        if (/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(id)) {
            return pnp.sp.web.lists.getById(id);
        } else {
            return pnp.sp.web.getList(id);
        }
    }
}
