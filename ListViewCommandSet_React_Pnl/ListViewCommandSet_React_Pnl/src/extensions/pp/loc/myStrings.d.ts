declare interface IPpCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'PpCommandSetStrings' {
  const strings: IPpCommandSetStrings;
  export = strings;
}
