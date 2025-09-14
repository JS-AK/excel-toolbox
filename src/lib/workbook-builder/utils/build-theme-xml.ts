import { XmlNode } from "../types/index.js";

import { XML_DECLARATION, XML_NAMESPACES } from "./constants.js";
import { buildXml } from "./build-xml.js";

export function buildThemeXml(): string {
	const theme: XmlNode = {
		attrs: {
			name: "Office Theme",
			"xmlns:a": XML_NAMESPACES.DRAWINGML,
		},
		children: [
			{
				children: [
					{
						attrs: { name: "Office" },
						children: [
							{
								children: [{ attrs: { lastClr: "000000", val: "windowText" }, tag: "a:sysClr" }],
								tag: "a:dk1",
							},
							{
								children: [{ attrs: { lastClr: "FFFFFF", val: "window" }, tag: "a:sysClr" }],
								tag: "a:lt1",
							},
							{
								children: [{ attrs: { val: "1F497D" }, tag: "a:srgbClr" }],
								tag: "a:dk2",
							},
							{
								children: [{ attrs: { val: "EEECE1" }, tag: "a:srgbClr" }],
								tag: "a:lt2",
							},
							{
								children: [{ attrs: { val: "4F81BD" }, tag: "a:srgbClr" }],
								tag: "a:accent1",
							},
							{
								children: [{ attrs: { val: "C0504D" }, tag: "a:srgbClr" }],
								tag: "a:accent2",
							},
							{
								children: [{ attrs: { val: "9BBB59" }, tag: "a:srgbClr" }],
								tag: "a:accent3",
							},
							{
								children: [{ attrs: { val: "8064A2" }, tag: "a:srgbClr" }],
								tag: "a:accent4",
							},
							{
								children: [{ attrs: { val: "4BACC6" }, tag: "a:srgbClr" }],
								tag: "a:accent5",
							},
							{
								children: [{ attrs: { val: "F79646" }, tag: "a:srgbClr" }],
								tag: "a:accent6",
							},
							{
								children: [{ attrs: { val: "0000FF" }, tag: "a:srgbClr" }],
								tag: "a:hlink",
							},
							{
								children: [{ attrs: { val: "800080" }, tag: "a:srgbClr" }],
								tag: "a:folHlink",
							},
						],
						tag: "a:clrScheme",
					},
					{
						attrs: { name: "Office" },
						children: [
							{
								children: [
									{ attrs: { typeface: "Calibri Light" }, tag: "a:latin" },
									{ attrs: { typeface: "" }, tag: "a:ea" },
									{ attrs: { typeface: "" }, tag: "a:cs" },
									{ attrs: { script: "Jpan", typeface: "游ゴシック Light" }, tag: "a:font" },
									{ attrs: { script: "Hang", typeface: "맑은 고딕" }, tag: "a:font" },
									{ attrs: { script: "Hans", typeface: "等线 Light" }, tag: "a:font" },
									{ attrs: { script: "Hant", typeface: "新細明體" }, tag: "a:font" },
									{ attrs: { script: "Arab", typeface: "Times New Roman" }, tag: "a:font" },
									{ attrs: { script: "Hebr", typeface: "Times New Roman" }, tag: "a:font" },
									{ attrs: { script: "Thai", typeface: "Tahoma" }, tag: "a:font" },
									{ attrs: { script: "Ethi", typeface: "Nyala" }, tag: "a:font" },
									{ attrs: { script: "Beng", typeface: "Vrinda" }, tag: "a:font" },
									{ attrs: { script: "Gujr", typeface: "Shruti" }, tag: "a:font" },
									{ attrs: { script: "Khmr", typeface: "MoolBoran" }, tag: "a:font" },
									{ attrs: { script: "Knda", typeface: "Tunga" }, tag: "a:font" },
									{ attrs: { script: "Guru", typeface: "Raavi" }, tag: "a:font" },
									{ attrs: { script: "Cans", typeface: "Euphemia" }, tag: "a:font" },
									{ attrs: { script: "Cher", typeface: "Plantagenet Cherokee" }, tag: "a:font" },
									{ attrs: { script: "Yiii", typeface: "Microsoft Yi Baiti" }, tag: "a:font" },
									{ attrs: { script: "Tibt", typeface: "Microsoft Himalaya" }, tag: "a:font" },
									{ attrs: { script: "Thaa", typeface: "MV Boli" }, tag: "a:font" },
									{ attrs: { script: "Deva", typeface: "Mangal" }, tag: "a:font" },
									{ attrs: { script: "Telu", typeface: "Gautami" }, tag: "a:font" },
									{ attrs: { script: "Taml", typeface: "Latha" }, tag: "a:font" },
									{ attrs: { script: "Syrc", typeface: "Estrangelo Edessa" }, tag: "a:font" },
									{ attrs: { script: "Orya", typeface: "Kalinga" }, tag: "a:font" },
									{ attrs: { script: "Mlym", typeface: "Kartika" }, tag: "a:font" },
									{ attrs: { script: "Laoo", typeface: "DokChampa" }, tag: "a:font" },
									{ attrs: { script: "Sinh", typeface: "Iskoola Pota" }, tag: "a:font" },
									{ attrs: { script: "Mong", typeface: "Mongolian Baiti" }, tag: "a:font" },
									{ attrs: { script: "Viet", typeface: "Times New Roman" }, tag: "a:font" },
									{ attrs: { script: "Uigh", typeface: "Microsoft Uighur" }, tag: "a:font" },
									{ attrs: { script: "Geor", typeface: "Sylfaen" }, tag: "a:font" },
								],
								tag: "a:majorFont",
							},
							{
								children: [
									{ attrs: { typeface: "Calibri" }, tag: "a:latin" },
									{ attrs: { typeface: "" }, tag: "a:ea" },
									{ attrs: { typeface: "" }, tag: "a:cs" },
									{ attrs: { script: "Jpan", typeface: "游ゴシック" }, tag: "a:font" },
									{ attrs: { script: "Hang", typeface: "맑은 고딕" }, tag: "a:font" },
									{ attrs: { script: "Hans", typeface: "等线" }, tag: "a:font" },
									{ attrs: { script: "Hant", typeface: "新細明體" }, tag: "a:font" },
									{ attrs: { script: "Arab", typeface: "Arial" }, tag: "a:font" },
									{ attrs: { script: "Hebr", typeface: "Arial" }, tag: "a:font" },
									{ attrs: { script: "Thai", typeface: "Tahoma" }, tag: "a:font" },
									{ attrs: { script: "Ethi", typeface: "Nyala" }, tag: "a:font" },
									{ attrs: { script: "Beng", typeface: "Vrinda" }, tag: "a:font" },
									{ attrs: { script: "Gujr", typeface: "Shruti" }, tag: "a:font" },
									{ attrs: { script: "Khmr", typeface: "DaunPenh" }, tag: "a:font" },
									{ attrs: { script: "Knda", typeface: "Tunga" }, tag: "a:font" },
									{ attrs: { script: "Guru", typeface: "Raavi" }, tag: "a:font" },
									{ attrs: { script: "Cans", typeface: "Euphemia" }, tag: "a:font" },
									{ attrs: { script: "Cher", typeface: "Plantagenet Cherokee" }, tag: "a:font" },
									{ attrs: { script: "Yiii", typeface: "Microsoft Yi Baiti" }, tag: "a:font" },
									{ attrs: { script: "Tibt", typeface: "Microsoft Himalaya" }, tag: "a:font" },
									{ attrs: { script: "Thaa", typeface: "MV Boli" }, tag: "a:font" },
									{ attrs: { script: "Deva", typeface: "Mangal" }, tag: "a:font" },
									{ attrs: { script: "Telu", typeface: "Gautami" }, tag: "a:font" },
									{ attrs: { script: "Taml", typeface: "Latha" }, tag: "a:font" },
									{ attrs: { script: "Syrc", typeface: "Estrangelo Edessa" }, tag: "a:font" },
									{ attrs: { script: "Orya", typeface: "Kalinga" }, tag: "a:font" },
									{ attrs: { script: "Mlym", typeface: "Kartika" }, tag: "a:font" },
									{ attrs: { script: "Laoo", typeface: "DokChampa" }, tag: "a:font" },
									{ attrs: { script: "Sinh", typeface: "Iskoola Pota" }, tag: "a:font" },
									{ attrs: { script: "Mong", typeface: "Mongolian Baiti" }, tag: "a:font" },
									{ attrs: { script: "Viet", typeface: "Arial" }, tag: "a:font" },
									{ attrs: { script: "Uigh", typeface: "Microsoft Uighur" }, tag: "a:font" },
									{ attrs: { script: "Geor", typeface: "Sylfaen" }, tag: "a:font" },
								],
								tag: "a:minorFont",
							},
						],
						tag: "a:fontScheme",
					},
					{
						attrs: { name: "Office" },
						children: [
							{
								children: [
									{
										children: [{ attrs: { val: "phClr" }, tag: "a:schemeClr" }],
										tag: "a:solidFill",
									},
									{
										attrs: { rotWithShape: "1" },
										children: [
											{
												children: [
													{
														attrs: { pos: "0" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{
																		attrs: { val: "50000" },
																		tag: "a:tint",
																	},
																	{
																		attrs: { val: "300000" },
																		tag: "a:satMod",
																	},
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
													{
														attrs: { pos: "35000" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{
																		attrs: { val: "37000" },
																		tag: "a:tint",
																	},
																	{
																		attrs: { val: "300000" },
																		tag: "a:satMod",
																	},
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
													{
														attrs: { pos: "100000" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{
																		attrs: { val: "15000" },
																		tag: "a:tint",
																	},
																	{
																		attrs: { val: "350000" },
																		tag: "a:satMod",
																	},
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
												],
												tag: "a:gsLst",
											},
											{
												attrs: { ang: "16200000", scaled: "1" },
												tag: "a:lin",
											},
										],
										tag: "a:gradFill",
									},
									{
										attrs: { rotWithShape: "1" },
										children: [
											{
												children: [
													{
														attrs: { pos: "0" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{
																		attrs: { val: "51000" },
																		tag: "a:shade",
																	},
																	{
																		attrs: { val: "130000" },
																		tag: "a:satMod",
																	},
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
													{
														attrs: { pos: "80000" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{
																		attrs: { val: "93000" },
																		tag: "a:shade",
																	},
																	{
																		attrs: { val: "130000" },
																		tag: "a:satMod",
																	},
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
													{
														attrs: { pos: "100000" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{
																		attrs: { val: "94000" },
																		tag: "a:shade",
																	},
																	{
																		attrs: { val: "135000" },
																		tag: "a:satMod",
																	},
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
												],
												tag: "a:gsLst",
											},
											{
												attrs: { ang: "16200000", scaled: "0" },
												tag: "a:lin",
											},
										],
										tag: "a:gradFill",
									},
								],
								tag: "a:fillStyleLst",
							},
							{
								children: [
									{
										attrs: { algn: "ctr", cap: "flat", cmpd: "sng", w: "9525" },
										children: [
											{
												children: [
													{
														attrs: { val: "phClr" },
														children: [
															{ attrs: { val: "95000" }, tag: "a:shade" },
															{ attrs: { val: "105000" }, tag: "a:satMod" },
														],
														tag: "a:schemeClr",
													},

												],
												tag: "a:solidFill",
											},
											{
												attrs: { val: "solid" },
												tag: "a:prstDash",
											},
										],
										tag: "a:ln",
									},
									{
										attrs: { algn: "ctr", cap: "flat", cmpd: "sng", w: "25400" },
										children: [
											{
												children: [
													{
														attrs: { val: "phClr" },
														tag: "a:schemeClr",
													},
												],
												tag: "a:solidFill",
											},
											{
												attrs: { val: "solid" },
												tag: "a:prstDash",
											},
										],
										tag: "a:ln",
									},
									{
										attrs: { algn: "ctr", cap: "flat", cmpd: "sng", w: "38100" },
										children: [
											{
												children: [
													{
														attrs: { val: "phClr" },
														tag: "a:schemeClr",
													},
												],
												tag: "a:solidFill",
											},
											{
												attrs: { val: "solid" },
												tag: "a:prstDash",
											},
										],
										tag: "a:ln",
									},
								],
								tag: "a:lnStyleLst",
							},
							{
								children: [
									{
										children: [
											{
												children: [
													{
														attrs: { blurRad: "40000", dir: "5400000", dist: "20000", rotWithShape: "0" },
														children: [
															{
																attrs: { val: "000000" },
																children: [{ attrs: { val: "38000" }, tag: "a:alpha" }],
																tag: "a:srgbClr",
															},
														],
														tag: "a:outerShdw",
													},
												],
												tag: "a:effectLst",
											},
										],
										tag: "a:effectStyle",
									},
									{
										children: [
											{
												children: [
													{
														attrs: { blurRad: "40000", dir: "5400000", dist: "23000", rotWithShape: "0" },
														children: [
															{
																attrs: { val: "000000" },
																children: [{ attrs: { val: "35000" }, tag: "a:alpha" }],
																tag: "a:srgbClr",
															},
														],
														tag: "a:outerShdw",
													},
												],
												tag: "a:effectLst",
											},
										],
										tag: "a:effectStyle",
									},
									{
										children: [
											{
												children: [
													{
														attrs: { blurRad: "40000", dir: "5400000", dist: "23000", rotWithShape: "0" },
														children: [
															{
																attrs: { val: "000000" },
																children: [{ attrs: { val: "35000" }, tag: "a:alpha" }],
																tag: "a:srgbClr",
															},
														],
														tag: "a:outerShdw",
													},
												],
												tag: "a:effectLst",
											},
											{
												children: [
													{
														attrs: { prst: "orthographicFront" },
														children: [{ attrs: { lat: "0", lon: "0", rev: "0" }, tag: "a:rot" }],
														tag: "a:camera",
													},
													{
														attrs: { dir: "t", rig: "threePt" },
														children: [{ attrs: { lat: "0", lon: "0", rev: "0" }, tag: "a:rot" }],
														tag: "a:lightRig",
													},
												],
												tag: "a:scene3d",
											},
											{
												children: [{ attrs: { h: "25400", w: "63500" }, tag: "a:bevelT" }],
												tag: "a:sp3d",
											},
										],
										tag: "a:effectStyle",
									},
								],
								tag: "a:effectStyleLst",
							},
							{
								children: [
									{
										children: [{ attrs: { val: "phClr" }, tag: "a:schemeClr" }],
										tag: "a:solidFill",
									},
									{
										attrs: { rotWithShape: "1" },
										children: [
											{
												children: [
													{
														attrs: { pos: "0" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{ attrs: { val: "40000" }, tag: "a:tint" },
																	{ attrs: { val: "350000" }, tag: "a:satMod" },
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
													{
														attrs: { pos: "40000" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{ attrs: { val: "45000" }, tag: "a:tint" },
																	{ attrs: { val: "99000" }, tag: "a:shade" },
																	{ attrs: { val: "350000" }, tag: "a:satMod" },
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
													{
														attrs: { pos: "100000" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{ attrs: { val: "20000" }, tag: "a:shade" },
																	{ attrs: { val: "255000" }, tag: "a:satMod" },
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
												],
												tag: "a:gsLst",
											},
											{
												attrs: { path: "circle" },
												children: [
													{
														attrs: { b: "180000", l: "50000", r: "50000", t: "-80000" },
														tag: "a:fillToRect",
													},
												],
												tag: "a:path",
											},
										],
										tag: "a:gradFill",
									},
									{
										attrs: { rotWithShape: "1" },
										children: [
											{
												children: [
													{
														attrs: { pos: "0" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{ attrs: { val: "80000" }, tag: "a:tint" },
																	{ attrs: { val: "300000" }, tag: "a:satMod" },
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
													{
														attrs: { pos: "100000" },
														children: [
															{
																attrs: { val: "phClr" },
																children: [
																	{ attrs: { val: "30000" }, tag: "a:shade" },
																	{ attrs: { val: "200000" }, tag: "a:satMod" },
																],
																tag: "a:schemeClr",
															},
														],
														tag: "a:gs",
													},
												],
												tag: "a:gsLst",
											},
											{
												attrs: { path: "circle" },
												children: [
													{
														attrs: { b: "50000", l: "50000", r: "50000", t: "50000" },
														tag: "a:fillToRect",
													},
												],
												tag: "a:path",
											},
										],
										tag: "a:gradFill",
									},
								],
								tag: "a:bgFillStyleLst",
							},
						],
						tag: "a:fmtScheme",
					},
				],
				tag: "a:themeElements",
			},
			{
				children: [],
				tag: "a:objectDefaults",
			},
			{
				children: [],
				tag: "a:extraClrSchemeLst",
			},
		],
		tag: "a:theme",
	};

	return [
		XML_DECLARATION,
		buildXml(theme),
	].join("\n");
}
