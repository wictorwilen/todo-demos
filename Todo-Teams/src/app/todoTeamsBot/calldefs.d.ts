export interface IncomingCallResponse {
  '@odata.type': string;
  value: IncomingCall[];
}

export interface IncomingCall {
  '@odata.type': string;
  changeType: string;
  resource: string;
  resourceData: ResourceData;
}

interface ResourceData {
  '@odata.type': string;
  state: string;
  direction: string;
  callbackUri: string;
  source: Source;
  targets: Target[];
  tenantId: string;
  myParticipantId: string;
  id: string;
  recordResourceLocation: string;
  recordResourceAccessToken: string;
}

interface Target {
  '@odata.type': string;
  identity: Identity2;
}

interface Identity2 {
  '@odata.type': string;
  application: Application;
}

interface Application {
  '@odata.type': string;
  id: string;
  tenantId?: any;
  identityProvider: string;
}

interface Source {
  '@odata.type': string;
  identity: Identity;
  region: string;
  languageId: string;
}

interface Identity {
  '@odata.type': string;
  user: User;
}

interface User {
  '@odata.type': string;
  id: string;
  tenantId: string;
  identityProvider: string;
}


interface ParticiapntsResponse {
  value: Value[];
}

interface Value {
  id: string;
  info: Info;
  isInLobby: boolean;
  isMuted: boolean;
  mediaStreams: MediaStream[];
  metadata: string;
}

interface MediaStream {
  sourceId: string;
  direction: string;
  label: string;
  mediaType: string;
  serverMuted: boolean;
}

interface Info {
  identity: Identity;
  languageId: string;
  region: string;
}

interface Identity {
  user: User;
}

interface User {
  id: string;
  tenantId: string;
  displayName: string;
}



