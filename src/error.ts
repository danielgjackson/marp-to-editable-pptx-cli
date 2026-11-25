/**
 * Error codes of {@link CLIError}. Each one is also the process exit status
 * the command-line tool uses for that kind of failure.
 */
export const CLIErrorCode = {
  /** A conversion failed, or an existing output was not overwritten. */
  GENERAL_ERROR: 1,
  /** Invalid usage: bad or unsupported options, or a missing input. */
  INVALID_OPTIONS: 2,
  /** The conversion was interrupted. */
  INTERRUPTED: 130,
} as const

export type CLIErrorCode = (typeof CLIErrorCode)[keyof typeof CLIErrorCode]

/** An error that stops the CLI, with the reason in `errorCode`. */
export class CLIError extends Error {
  readonly errorCode: CLIErrorCode

  constructor(message: string, errorCode: CLIErrorCode = CLIErrorCode.GENERAL_ERROR) {
    super(message)
    this.name = 'CLIError'
    this.errorCode = errorCode
  }
}
