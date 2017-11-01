import * as pnp from 'sp-pnp-js';
import * as chalk from 'chalk';
import { List } from 'sp-pnp-js';
import { IViewCommandOptions, IAddCommandOptions, IUpdateCommandOptions, IDeleteCommandOptions, IListCommandOption } from './params/index';

export class SubscriptionManager {

    public async viewSubscriptions(opts: IViewCommandOptions, verbose = true): Promise<any> {
        let list = this.getList(opts.list);
        let data: any;
        if (opts.id) {
            data = await list.subscriptions.getById(opts.id).get();
        } else {
            data = await list.subscriptions.get();
        }

        if (verbose) {
            console.log('Subscription data received:');
            console.log(data);
        }

        return data;
    }

    public async addSubscription(opts: IAddCommandOptions): Promise<any> {
        if (!opts.exp) {
            opts.exp = this.getDefaultExpiration();
        }

        let list = this.getList(opts.list);
        let data = await list.subscriptions.add(opts.url, opts.exp);
        console.log('Subscription added:');
        console.log(data);

        return data;
    }

    public async deleteSubscription(opts: IDeleteCommandOptions): Promise<any> {
        let list = this.getList(opts.list);
        await list.subscriptions.getById(opts.id).delete();
        console.log(`Subscription ${opts.id} deleted`);
    }

    public async deleteAllSubscriptions(opts: IListCommandOption): Promise<any> {

        let subscriptions = await this.viewSubscriptions(opts, false);
        let deletePromises: any[] = [];
        subscriptions.forEach((subscription: any) => {
            deletePromises.push(this.deleteSubscription({
                id: subscription.id,
                list: opts.list
            }));
        });

        return Promise.all(deletePromises);
    }

    public async updateSubscription(opts: IUpdateCommandOptions): Promise<any> {
        if (!opts.exp) {
            opts.exp = this.getDefaultExpiration();
        }
        let list = this.getList(opts.list);
        await list.subscriptions.getById(opts.id).update(opts.exp);

        console.log(`Subscription ${opts.id} updated`);
    }

    private getDefaultExpiration(): string {
        let now = new Date();
        now.setMonth(now.getMonth() + 5);
        now.setHours(24 * 25);
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
