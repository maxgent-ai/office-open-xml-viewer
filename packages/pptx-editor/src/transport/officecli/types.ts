import type { NonEmptyReadonlyArray } from '../../domain/command';
import type {
  OFFICECLI_BATCH_SCHEMA_VERSION,
  OFFICECLI_COMMAND_TYPES,
  OFFICECLI_ELEMENT_TYPES,
  OFFICECLI_VERSION,
} from './constants';

export type OfficeCliCommandType =
  (typeof OFFICECLI_COMMAND_TYPES)[keyof typeof OFFICECLI_COMMAND_TYPES];

export type OfficeCliProps = Readonly<Record<string, string>>;

export interface OfficeCliAddCommand {
  readonly command: typeof OFFICECLI_COMMAND_TYPES.ADD;
  readonly parent: string;
  readonly type: typeof OFFICECLI_ELEMENT_TYPES.SHAPE;
  readonly props: OfficeCliProps;
}

export interface OfficeCliSetCommand {
  readonly command: typeof OFFICECLI_COMMAND_TYPES.SET;
  readonly path: string;
  readonly props: OfficeCliProps;
}

export interface OfficeCliRemoveCommand {
  readonly command: typeof OFFICECLI_COMMAND_TYPES.REMOVE;
  readonly path: string;
}

export type OfficeCliCommand =
  | OfficeCliAddCommand
  | OfficeCliSetCommand
  | OfficeCliRemoveCommand;

/** Product envelope around the native command array accepted by `officecli batch`. */
export interface OfficeCliBatch {
  readonly schemaVersion: typeof OFFICECLI_BATCH_SCHEMA_VERSION;
  readonly officecliVersion: typeof OFFICECLI_VERSION;
  readonly commandId: string;
  readonly commands: NonEmptyReadonlyArray<OfficeCliCommand>;
}
