declare module "node:fs" {
  export function appendFileSync(path: string, data: string, options?: { mode?: number } | string): void;
  export function copyFileSync(src: string, dest: string): void;
  export function existsSync(path: string): boolean;
  export function mkdirSync(path: string, options?: { recursive?: boolean; mode?: number }): void;
  export function readFileSync(path: string, encoding: "utf8"): string;
  export function readdirSync(path: string): string[];
  export function renameSync(oldPath: string, newPath: string): void;
  export function statSync(path: string): { isDirectory(): boolean; isFile(): boolean };
  export function unlinkSync(path: string): void;
  export function writeFileSync(path: string, data: string, options?: { mode?: number } | string): void;
}

declare module "node:child_process" {
  export function spawnSync(command: string, args?: readonly string[], options?: { encoding?: "utf8"; maxBuffer?: number }): { status: number | null; stdout?: string; stderr?: string; error?: unknown };
}

declare module "node:crypto" {
  export function createHash(algorithm: string): { update(data: string): { digest(encoding: "hex"): string } };
}
