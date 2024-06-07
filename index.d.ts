declare module 'node-activex' {
  export class Object {
    constructor(id: string, opt?: any);
  }
}

declare global {
  function ActiveXObject(id: string, opt?: any): any;
}
