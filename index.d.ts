export class Object {
  constructor(id: string, options?: ActiveXObjectOptions);
  
  // Define properties typically found on COM objects
  __id?: string;
  __value?: any;
  __type?: any[];
  __methods?: string[];
  __vars?: string[];

  // Define general method signatures if any known
  [key: string]: any;
}

/** @deprecated Use `ActiveXObjectOptions` instead. */
export interface ActiveXOptions extends ActiveXObjectOptions {}

export class Variant {
  constructor(value?: any, type?: VariantType);
  assign(value: any): void;
  cast(type: VariantType): void;
  clear(): void;
  valueOf(): any;
}

export type VariantType =
  | 'int' | 'uint' | 'int8' | 'char' | 'uint8' | 'uchar' | 'byte'
  | 'int16' | 'short' | 'uint16' | 'ushort'
  | 'int32' | 'uint32'
  | 'int64' | 'long' | 'uint64' | 'ulong'
  | 'currency' | 'float' | 'double' | 'string'
  | 'date' | 'decimal' | 'variant' | 'null' | 'empty'
  | 'byref' | 'pbyref';

export function cast(value: any, type: VariantType): any;

// Utility function to release COM objects
export function release(...objects: any[]): void;

declare global {
  function ActiveXObject(id: string, options?: ActiveXObjectOptions): any;
  function ActiveXObject(obj: Record<string, any>): any;

  interface ActiveXObjectOptions {
    /** Allow activating existing object instance. */
    activate?: boolean;
    
    /** Allow using the name of the file in the ROT. **/
    getobject?: boolean;

    /** Allow using type information. */
    type?: boolean;
  }
}
