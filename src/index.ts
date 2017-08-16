#!/usr/bin/env node

import * as minimist from 'minimist';
import * as chalk from 'chalk';
import * as inquirer from 'inquirer';
import * as spauth from 'node-sp-auth';
import { SubscriptionManager } from './SubscriptionManager';
import { PnPConfig } from './PnPConfig';

declare var global: any;

let Preferences: any = require('preferences');
let args = minimist(process.argv.slice(2));
let manager = new SubscriptionManager(args);
let prefs = new Preferences('sphooks', {
    siteUrl: ''
});

function printError(error: any): void {
    console.log(chalk.red(error));
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

if (args._.indexOf('login') !== -1) {
    login();
} else {
    if (!prefs.siteUrl) {
        console.log(chalk.red('Use \'sphooks login\' command first'));
        process.exit();
    }

    PnPConfig.configure(prefs.siteUrl);

    console.log('Connecting to site: ' + prefs.siteUrl);
    console.log(chalk.green('Connected'));

    if (args._.indexOf('view') !== -1) {
        manager.viewSubscriptions().catch(printError);
    } else if (args._.indexOf('add') !== -1) {
        manager.addSubscription().catch(printError);
    } else if (args._.indexOf('delete') !== -1) {
        manager.deleteSubscription().catch(printError);
    } else if (args._.indexOf('deleteAll') !== -1) {
        manager.deleteAllSubscriptions().catch(printError);
    } else if (args._.indexOf('update') !== -1) {
        manager.updateSubscription();
    } else {
        console.log(chalk.red('Wrong parameters'));
    }
}
