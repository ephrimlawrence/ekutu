export type ValueType = string | number | Date | unknown | Array<any>;

export class HeaderRow {
    /**
     * The title of the header (as will be shown in excel)
     */
    title!: string;

    id!: string;
    type?: ValueType;
    sub?: HeaderRow | HeaderRow[];
    hidden: boolean = false;
}
