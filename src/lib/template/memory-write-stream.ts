export class MemoryWriteStream {
	private chunks: Buffer[] = [];

	write(chunk: string | Buffer): boolean {
		this.chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk, "utf-8"));
		return true;
	}

	end(): void {
		// no-op
	}

	toBuffer(): Buffer {
		return Buffer.concat(this.chunks);
	}
}
