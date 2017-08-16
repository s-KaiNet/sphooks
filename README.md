# sphooks - cross-platform CLI for managing SharePoint list web hooks

[![NPM](https://nodei.co/npm/sphooks.png?mini=true)](https://nodei.co/npm/sphooks/)
[![npm version](https://badge.fury.io/js/sphooks.svg)](https://badge.fury.io/js/sphooks)

### Need help on SharePoint with Node.JS? Join our gitter chat and ask question! [![Gitter chat](https://badges.gitter.im/gitterHQ/gitter.png)](https://gitter.im/sharepoint-node/Lobby)

Built with Node.JS, Typescript, [sp-pnp-js](https://github.com/SharePoint/PnP-JS-Core) and [node-sp-auth](https://github.com/s-KaiNet/node-sp-auth) 

## Prerequisites
 - Node.JS >= 6.x


## Installation

```bash
npm install sphooks -g
```

## Usage

#### Loign:
```
sphooks login
```
Runs interactive web login session and allows you to authenticate against SharePoint. Authentication stored in a file in an encrypted manner. If you want to change the site, run `sphooks login` once again. 


#### View:

```
sphooks view --list <list id or url> [--id] <subscription id>
```
 - `--list` - required, GUID or list server relative url, i.e. "sites/dev/My List"
 - `--id` - optional GUID, subscription id. When omitted, all subscriptions will be returned

 Sample:   
 ```
 sphooks view --list "sites/dev/Shared Documents" --id ce38389e-a91d-4df4-b924-0b1956b4640e
```

#### Add: 
```
sphooks add --list <list id or url> --url <notification url> [--exp] <expiration date>
```
 - `--list` - required, GUID or list server relative url, i.e. "sites/dev/My List"
 - `--url` - required string, notification url
 - `--exp` - optional date as ISO formatted string, i.e. '2017-08-16T16:40:09.189Z'. When omitted, date.now + 6 months will be used (maximum allowed for SharePoint web hook)

 Sample:  
 ```
 sphooks add --list 38e3ec1c-1bad-42fa-a60a-2a6c1e49cfba --url https://myfunc.azurewebsites.com/api/webhooks/ --exp 2017-08-16T16:40:09.189Z
 ```


#### Update:
```
sphooks update --list <list id or url> --id <subscription id> [--exp] <expiration date>
```

 - `--list` - required, GUID or list server relative url, i.e. "sites/dev/My List"
  - `--id` - required GUID, subscription id
 - `--exp` - optional date as ISO formatted string, i.e. '2017-08-16T16:40:09.189Z'. When omitted, date.now + 6 months will be used (maximum allowed for SharePoint web hook)

 Sample:  
 ```
 sphooks update --list "sites/dev/Shared Documents" --url https://myfunc.azurewebsites.com/api/webhooks/ --id 183cdbd9-446d-4ff3-a9d4-f01925f55022 --exp 2017-08-16T16:40:09.189Z
```

#### Delete:
```
sphooks delete --list <list id or url> --id <subscription id>
```

 - `--list` - required, GUID or list server relative url, i.e. "sites/dev/My List"
 - `--id` - required GUID, subscription id

Sample 
```
sphooks delete --list "sites/dev/Shared Documents" --id 183cdbd9-446d-4ff3-a9d4-f01925f55022
```

 #### Delete all:
 ```
 sphooks deleteAll --list <list id or url>
 ```
  - `--list` - required, GUID or list server relative url, i.e. "sites/dev/My List"

Sample 
```
sphooks deleteAll --list "sites/dev/Shared Documents" 
```
