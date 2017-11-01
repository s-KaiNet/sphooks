export interface IViewCommandOptions extends IListCommandOption {
    id?: string;
}

export interface IAddCommandOptions extends IListCommandOption {
    url: string;
    exp?: string;
}

export interface IUpdateCommandOptions extends IListCommandOption {
    id: string;
    exp?: string;
}

export interface IDeleteCommandOptions extends IListCommandOption {
    id?: string;
}

export interface IListCommandOption {
    list: string;
}
