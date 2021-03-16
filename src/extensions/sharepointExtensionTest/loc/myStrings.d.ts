declare interface ISharepointExtensionTestCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SharepointExtensionTestCommandSetStrings' {
  const strings: ISharepointExtensionTestCommandSetStrings;
  export = strings;
}
