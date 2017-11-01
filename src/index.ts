#!/usr/bin/env node

import * as chalk from 'chalk';
import * as program from 'commander';
import * as path from 'path';
import * as inquirer from 'inquirer';
import * as spauth from 'node-sp-auth';
import { SubscriptionManager } from './SubscriptionManager';
import { PnPConfig } from './PnPConfig';
import {
    IViewCommandOptions,
    IAddCommandOptions,
    IUpdateCommandOptions,
    IDeleteCommandOptions,
    IListCommandOption
} from './params/index';

const { version } = require(path.join(__dirname, '..', 'package.json'));

let manager = new SubscriptionManager();
let Preferences: any = require('preferences');
let prefs = new Preferences('sphooks', {
    siteUrl: ''
});

program
    .version(version)
    .name('sphooks')
    .usage('[command]')
    .description('Command line utility for managing SharePoint list web hooks ')
    .command('login')
    .description('Log you in into the SharePoint using interactive web session')
    .action(login);

program
    .command('view')
    .description('Prints webhooks registered for the list')
    .option('--list <list id or url>', 'required, GUID or list server relative url, i.e. "sites/dev/My List"')
    .option('--id <guid>', 'optional GUID, subscription id. When omitted, all subscriptions will be returned')
    .action((opts: IViewCommandOptions) => {
        setupAction(() => {
            assert(opts.list, '--list');

            manager.viewSubscriptions(opts).catch(printError);
        });
    });

program
    .command('add')
    .description('Adds webhook to a list')
    .option('--list <list id or url>', 'required, GUID or list server relative url, i.e. "sites/dev/My List"')
    .option('--url <url>', 'required string, notification url')
    .option('--exp <iso date>', 'optional date as ISO formatted string, i.e. \'2017-08-16T16:40:09.189Z\'. When omitted, date.now + 6 months will be used (maximum allowed for SharePoint web hook)')
    .action((opts: IAddCommandOptions) => {
        setupAction(() => {
            assert(opts.list, '--list');
            assert(opts.url, '--url');

            manager.addSubscription(opts).catch(printError);
        });
    });

program
    .command('update')
    .description('Updates expiration date for registered webhook')
    .option('--list <list id or url>', 'required, GUID or list server relative url, i.e. "sites/dev/My List"')
    .option('--id <id>', 'required GUID, subscription id')
    .option('--exp <iso date>', 'optional date as ISO formatted string, i.e. \'2017-08-16T16:40:09.189Z\'. When omitted, date.now + 6 months will be used (maximum allowed for SharePoint web hook)')
    .action((opts: IUpdateCommandOptions) => {
        setupAction(() => {
            assert(opts.list, '--list');
            assert(opts.id, '--id');

            manager.updateSubscription(opts).catch(printError);
        });
    });

program
    .command('delete')
    .description('Deletes webhook by its id')
    .option('--list <list id or url>', 'required, GUID or list server relative url, i.e. "sites/dev/My List"')
    .option('--id <id>', 'required GUID, subscription id')
    .action((opts: IDeleteCommandOptions) => {
        setupAction(() => {
            assert(opts.list, '--list');
            assert(opts.id, '--id');

            manager.deleteSubscription(opts).catch(printError);
        });
    });

program
    .command('deleteAll')
    .description('Deletes all webhooks for the list')
    .option('--list <list id or url>', 'required, GUID or list server relative url, i.e. "sites/dev/My List"')
    .action((opts: IListCommandOption) => {
        setupAction(() => {
            assert(opts.list, '--list');

            manager.deleteAllSubscriptions(opts).catch(printError);
        });
    });

program.parse(process.argv);

if (program.args.length === 0) {
    program.help();
}

function setupAction(cb: () => void) {
    setup();
    cb();
}

function printError(error: any): void {
    console.log(chalk.red(error));
}

function assert(value: any, name: string) {
    if (typeof value === 'undefined') {
        console.log(chalk.red(`${name} - parameter is required`));
        process.exit();
    }
}

function login() {
    let questions = [
        {
            name: 'siteUrl',
            type: 'input',
            message: 'Enter your SharePoint site url:',
            validate: (value: string) => {
                if (value.length) {
                    return true;
                } else {
                    return 'Please enter your site url';
                }
            }
        }
    ];

    inquirer.prompt(questions)
        .then(answers => {
            prefs.siteUrl = answers.siteUrl;
            return spauth.getAuth(prefs.siteUrl, {
                ondemand: true
            });
        })
        .then(() => {
            process.exit();
        })
        .catch(err => {
            console.log(chalk.red(err));
            process.exit();
        });
}

function setup() {
    if (!prefs.siteUrl) {
        console.log(chalk.red('Use \'sphooks login\' command first'));
        process.exit();
    }

    PnPConfig.configure(prefs.siteUrl);

    console.log('Connecting to site: ' + prefs.siteUrl);
    console.log(chalk.green('Connected'));
}
