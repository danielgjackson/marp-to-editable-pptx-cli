// Type declarations for the programmatic API (see "API" in README.md).
// Kept in step with src/cli.ts and src/error.ts by hand.

/**
 * Error codes of {@link CLIError}. Each one is also the process exit status
 * the command-line tool uses for that kind of failure.
 */
export declare const CLIErrorCode: {
  /** A conversion failed, or an existing output was not overwritten. */
  readonly GENERAL_ERROR: 1
  /** Invalid usage: bad or unsupported options, or a missing input. */
  readonly INVALID_OPTIONS: 2
  /** The conversion was interrupted. */
  readonly INTERRUPTED: 130
}
export type CLIErrorCode = (typeof CLIErrorCode)[keyof typeof CLIErrorCode]

/** An error that stops the CLI, with the reason in `errorCode`. */
export declare class CLIError extends Error {
  readonly errorCode: CLIErrorCode
  constructor(message: string, errorCode?: CLIErrorCode)
}

/** Options for {@link marpCli}. */
export interface MarpCLIAPIOptions {
  /**
   * Read Markdown from standard input when no input is given, as the
   * command-line tool does when data is piped to it. Default: `false`. An
   * explicit `-` input always reads stdin.
   */
  stdin?: boolean
}

/**
 * Run the CLI with the given arguments, as Marp CLI's `marpCli()` does.
 *
 * Resolves with the exit status the command would have returned (0 success,
 * 1 a conversion failed or an output was not overwritten, 130 interrupted).
 * Rejects with a {@link CLIError} for invalid usage, and with other errors
 * for unexpected failures.
 */
export declare function marpCli(argv: string[], options?: MarpCLIAPIOptions): Promise<number>
