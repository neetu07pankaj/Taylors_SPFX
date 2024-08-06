declare interface IRiskOwnerCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'RiskOwnerCommandSetStrings' {
  const strings: IRiskOwnerCommandSetStrings;
  export = strings;
}
