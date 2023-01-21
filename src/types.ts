export type ValueType = string | number | Date | unknown | Array<any>;

export class HeaderRow {
    /**
     * The title of the header (as will be shown in excel)
     */
    title!: string;

    id!: string;
    path!: string;
    type?: ValueType;
    sub?: HeaderRow | HeaderRow[];
    hidden: boolean = false;

   get isNestedPath() {
        return this.path.includes(".");
    }

    getValueByPath(item: any, route: string): any {
        const path = route.split(".");
        let value = item;

        for (const p of path) {
            value = value[p];

            if(value == null) return null;
        }

        return value;
    }
}
