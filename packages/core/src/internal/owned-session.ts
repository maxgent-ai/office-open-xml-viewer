/** Minimal structural contract for a resource-owning session. */
export interface ClosableOwnedSession {
  close(): Promise<void>;
}

/**
 * Run one operation against a newly opened owned session and close exactly once.
 * An operation failure is primary; cleanup is surfaced only when the operation
 * itself succeeded. This is lifecycle mechanics, not a shared format state
 * machine.
 */
export async function usingOwnedSession<TSession extends ClosableOwnedSession, TResult>(
  open: () => Promise<TSession>,
  operation: (session: TSession) => Promise<TResult>,
): Promise<TResult> {
  const session = await open();
  let operationError: unknown;
  try {
    return await operation(session);
  } catch (error) {
    operationError = error;
    throw error;
  } finally {
    try {
      await session.close();
    } catch (cleanupError) {
      if (operationError === undefined) throw cleanupError;
    }
  }
}
