import { describe, expect, it } from "vitest";

import { escapeXml } from "./escape-xml.js";

describe("escapeXml", () => {
	it("should escape special XML characters", () => {
		expect(escapeXml("&")).toBe("&amp;");
		expect(escapeXml("<")).toBe("&lt;");
		expect(escapeXml(">")).toBe("&gt;");
		expect(escapeXml("\"")).toBe("&quot;");
		expect(escapeXml("'")).toBe("&apos;");
	});

	it("should handle strings with multiple special characters", () => {
		expect(escapeXml("<div class=\"test\">")).toBe("&lt;div class=&quot;test&quot;&gt;");
		expect(escapeXml("John & Jane's Code")).toBe("John &amp; Jane&apos;s Code");
		expect(escapeXml("1 < 2 && 2 > 1")).toBe("1 &lt; 2 &amp;&amp; 2 &gt; 1");
	});

	it("should handle strings with no special characters", () => {
		expect(escapeXml("Hello World")).toBe("Hello World");
		expect(escapeXml("12345")).toBe("12345");
		expect(escapeXml("普通のテキスト")).toBe("普通のテキスト");
	});

	it("should handle empty strings", () => {
		expect(escapeXml("")).toBe("");
	});

	it("should handle already escaped strings", () => {
		expect(escapeXml("&amp;")).toBe("&amp;amp;");
		expect(escapeXml("&lt;div&gt;")).toBe("&amp;lt;div&amp;gt;");
		expect(escapeXml("&quot;test&quot;")).toBe("&amp;quot;test&amp;quot;");
	});

	it("should handle mixed content", () => {
		expect(escapeXml("Price: $10 < $20 & \"sale\"")).toBe("Price: $10 &lt; $20 &amp; &quot;sale&quot;");
		expect(escapeXml("Don't forget!")).toBe("Don&apos;t forget!");
		expect(escapeXml("A<B && B>C")).toBe("A&lt;B &amp;&amp; B&gt;C");
	});

	it("should handle special cases", () => {
		expect(escapeXml("<<<>>>")).toBe("&lt;&lt;&lt;&gt;&gt;&gt;");
		expect(escapeXml("\"\"''\"\"")).toBe("&quot;&quot;&apos;&apos;&quot;&quot;");
		expect(escapeXml("&&&&")).toBe("&amp;&amp;&amp;&amp;");
	});

	it("should handle complex strings", () => {
		const input = "The <b>quick</b> brown & fox 'jumps' over the \"lazy\" dog > 5 times";
		const expected = "The &lt;b&gt;quick&lt;/b&gt; brown &amp; fox &apos;jumps&apos; over the &quot;lazy&quot; dog &gt; 5 times";
		expect(escapeXml(input)).toBe(expected);
	});
});
