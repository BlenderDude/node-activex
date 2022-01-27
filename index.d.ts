declare module "winax" {
    export type VariantType = 
        | 'int' // VT_INT
        | 'uint' // VT_UINT
        | 'int8' // VT_I1
        | 'char' // VT_I1
        | 'uint8' // VT_UI1
        | 'uchar' // VT_UI1
        | 'byte' // VT_UI1
        | 'int16' // VT_I2
        | 'short' // VT_I2
        | 'uint16' // VT_UI2
        | 'ushort' // VT_UI2
        | 'int32' // VT_I4
        | 'uint32' // VT_UI4
        | 'int64' // VT_I8
        | 'long' // VT_I8
        | 'uint64' // VT_UI8
        | 'ulong' // VT_UI8
        | 'currency' // VT_CY
        | 'float' // VT_R4
        | 'double' // VT_R8
        | 'string' // VT_BSTR
        | 'date' // VT_DATE
        | 'decimal' // VT_DECIMAL
        | 'variant' // VT_VARIANT
        | 'null' // VT_NULL
        | 'empty' // VT_EMPTY
        | 'bool' // VT_BOOL
        | 'byref' // VT_BYREF or use prefix `p` to indicate reference to the current type

    export type VariantValue = string | boolean | number | Variant | VariantValue[];
    
    export class Variant {
        constructor(value: VariantValue, type?: VariantType);
        assign(value: number | string): void;
        cast(type: VariantType): void;
        clear(): void;
        __id: string;
        __value: any;
        __type: any;
        __methods: any;
        __vars: any;
    }

    export interface ObjectOptions {
        activate: boolean;  // Allow activate existence object instance, false by default
        async: boolean; // Allow asynchronous calls, true by default (not implemented, for future usage)
        type: boolean; // Allow using type information, true by default
    }

    export type ComObject = Record<string, any>;

    export const Object: {
        new <T = any>(cls: string, options?: ObjectOptions): T;
        new <T = any>(options: ComObject): T;
    };
    export function release(...args: any[]): void;
    export function cast(value: VariantValue, type: VariantType): Variant;
}