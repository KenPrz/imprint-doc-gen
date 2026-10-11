export class RenderError extends Error {
  readonly code: string;
  readonly part?: string;
  constructor(code: string, message: string, part?: string) {
    super(message); this.name = 'RenderError'; this.code = code; this.part = part;
  }
}
