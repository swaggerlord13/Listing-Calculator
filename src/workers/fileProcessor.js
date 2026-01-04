// eslint-disable-next-line no-redeclare
/* global self, importScripts, XLSX */

// Load XLSX library
importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");

// ============================================================================
// OPTIMIZED CATEGORY MATCHING SYSTEM
// ============================================================================

const STOPWORDS = new Set([
  "with",
  "for",
  "and",
  "or",
  "the",
  "a",
  "an",
  "in",
  "on",
  "at",
  "to",
  "from",
  "by",
  "of",
  "is",
  "was",
  "are",
  "were",
  "been",
]);

class CategoryIndex {
  constructor() {
    this.tokenIndex = new Map();
    this.categories = [];
    this.isInitialized = false;
  }

  tokenize(text) {
    return text
      .toLowerCase()
      .replace(/rrp\s*£\d+/gi, "")
      .replace(/[^\w\s]/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(/\s+/)
      .filter((word) => word.length > 2 && !STOPWORDS.has(word));
  }

  build(categoryMap) {
    if (!categoryMap || categoryMap.length === 0) {
      console.warn("Empty category map provided");
      return;
    }

    const startTime = Date.now();
    this.tokenIndex.clear();
    this.categories = [];

    categoryMap.forEach((cat, idx) => {
      const categoryId = cat["Category ID"] || cat.CategoryID || cat.categoryId;
      const categoryPath =
        cat["Category Path"] || cat.CategoryPath || cat.categoryPath || "";

      if (!categoryId || !categoryPath) return;

      const depth = categoryPath.split(">").length;
      const pathTokens = this.tokenize(categoryPath);

      this.categories.push({
        id: categoryId,
        path: categoryPath,
        tokens: pathTokens,
        depth: depth,
        index: idx,
      });

      pathTokens.forEach((token) => {
        if (!this.tokenIndex.has(token)) {
          this.tokenIndex.set(token, new Set());
        }
        this.tokenIndex.get(token).add(idx);
      });
    });

    this.isInitialized = true;

    const elapsed = Date.now() - startTime;

    // Notify main thread about cache status
    self.postMessage({
      type: "category_cache",
      count: this.categories.length,
      buildTime: elapsed,
      isNewUpload: true,
    });
  }

  match(title) {
    if (!this.isInitialized || this.categories.length === 0) {
      return { categoryId: 47155, categoryPath: "Default", score: 0 };
    }

    const titleTokens = this.tokenize(title);

    if (titleTokens.length === 0) {
      return { categoryId: 47155, categoryPath: "Default", score: 0 };
    }

    const scores = new Map();

    titleTokens.forEach((titleToken) => {
      const matchingCategories = this.tokenIndex.get(titleToken);

      if (!matchingCategories) return;

      matchingCategories.forEach((categoryIndex) => {
        const category = this.categories[categoryIndex];

        if (!scores.has(categoryIndex)) {
          scores.set(categoryIndex, {
            exact: 0,
            partial: 0,
            depth: category.depth,
            category: category,
          });
        }

        const scoreData = scores.get(categoryIndex);

        if (category.tokens.includes(titleToken)) {
          scoreData.exact += 10;
        }

        category.tokens.forEach((catToken) => {
          if (catToken.includes(titleToken) || titleToken.includes(catToken)) {
            scoreData.partial += 5;
          }
        });
      });
    });

    let bestMatch = null;
    let bestScore = -1;

    scores.forEach((scoreData) => {
      const exactScore = scoreData.exact;
      const partialScore = scoreData.partial;
      const depthBonus = scoreData.depth * 3;

      const matchedTokens = new Set();
      titleTokens.forEach((tt) => {
        if (
          scoreData.category.tokens.some(
            (ct) => ct === tt || ct.includes(tt) || tt.includes(ct)
          )
        ) {
          matchedTokens.add(tt);
        }
      });
      const coverageBonus = matchedTokens.size * 8;

      const totalScore = exactScore + partialScore + depthBonus + coverageBonus;

      if (totalScore > bestScore) {
        bestScore = totalScore;
        bestMatch = scoreData.category;
      }
    });

    if (!bestMatch || bestScore < 5) {
      return { categoryId: 47155, categoryPath: "Default", score: 0 };
    }

    return {
      categoryId: bestMatch.id,
      categoryPath: bestMatch.path,
      score: bestScore,
    };
  }
}

const categoryIndex = new CategoryIndex();

function matchCategory(title) {
  return categoryIndex.match(title).categoryId;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const postageRateTable = [
  { maxWeight: 0, postage: 1.9662, code: 2 },
  { maxWeight: 2, postage: 3.8136, code: 5 },
  { maxWeight: 5, postage: 4.156, code: 10 },
  { maxWeight: 10, postage: 3.656, code: 10 },
];

const round = (num, decimals = 2) => {
  return Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals);
};

const getExcelColumn = (col) => {
  let temp,
    letter = "";
  while (col >= 0) {
    temp = col % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
};

const formatDateToSKU = (dateStr) => {
  const [year, month, day] = dateStr.split("-");
  return `${day}${month}${year.slice(2)}`;
};

const getPostageInfo = (weight, postageRates = null) => {
  const rates = postageRates || postageRateTable;
  let selectedRate = rates[0];
  for (let i = 0; i < rates.length; i++) {
    if (weight >= rates[i].maxWeight) {
      selectedRate = rates[i];
    } else {
      break;
    }
  }
  return { postage: selectedRate.postage, code: selectedRate.code };
};

const generateTags = (title) => {
  if (!title) return "";
  const words = title
    .toLowerCase()
    .replace(/[^\w\s,]/g, "")
    .split(/[\s,]+/)
    .filter((w) => w.length > 2);
  return words.slice(0, 15).join(", ");
};

// Extract item specifics from title and description
const extractItemSpecifics = (title, description) => {
  const specifics = {};
  const text = `${title} ${description}`.toLowerCase();

  // Size patterns (UK sizes, EU sizes, US sizes, general sizes)
  const sizeMatch = text.match(/\b(size[:\s]*)?(\d+(\.\d+)?)\s*(uk|eu|us|cm|mm|ml|l|kg|g|oz|inches?|in)?\b/i) ||
                   text.match(/\b(small|medium|large|x-?large|xx-?large|xs|s|m|l|xl|xxl|xxxl)\b/i);
  if (sizeMatch) {
    specifics.Size = sizeMatch[0].trim();
  }

  // Color patterns
  const colorMatch = text.match(/\b(colour?[:\s]*)?(black|white|red|blue|green|yellow|pink|purple|orange|grey|gray|brown|beige|navy|gold|silver|multi-?colou?r)\b/i);
  if (colorMatch) {
    specifics.Color = colorMatch[2] || colorMatch[0];
  }

  // Material patterns
  const materialMatch = text.match(/\b(material[:\s]*)?(cotton|polyester|leather|wool|silk|denim|suede|nylon|plastic|metal|wood|glass|ceramic|rubber)\b/i);
  if (materialMatch) {
    specifics.Material = materialMatch[2] || materialMatch[0];
  }

  // Gender patterns
  const genderMatch = text.match(/\b(men'?s?|women'?s?|unisex|boys?|girls?|kids?|children'?s?)\b/i);
  if (genderMatch) {
    specifics.Gender = genderMatch[0];
  }

  // Age group patterns
  const ageMatch = text.match(/\b(adult|child|baby|toddler|infant|teen)\b/i);
  if (ageMatch) {
    specifics.AgeGroup = ageMatch[0];
  }

  return specifics;
};

const shortenTitle = (title, maxLength = 70) => {
  if (!title) return "";
  const fillers = [
    "with",
    "for",
    "the",
    "and",
    "&",
    "-",
    "a",
    "an",
    "featuring",
    "includes",
    "comes with",
    "perfect for",
    "that",
  ];
  const words = title.split(/[\s,]+/);
  let result = [];
  let length = 0;

  for (let word of words) {
    const cleanWord = word.replace(/[^\w\s]/g, "");
    if (fillers.includes(cleanWord.toLowerCase())) continue;
    if (length + word.length + 1 <= maxLength) {
      result.push(word);
      length += word.length + 1;
    } else {
      break;
    }
  }
  return result.join(" ").substring(0, maxLength).trim();
};

const parsePostageRates = (postageData) => {
  const rates = [];
  for (let i = 1; i < postageData.length; i++) {
    const row = postageData[i];
    if (!row || row.length === 0) continue;

    const maxWeight = parseFloat(row[0]) || 0;
    const postage = parseFloat(row[1]) || 0;
    const code = parseInt(row[6]) || 0;

    if (postage > 0) {
      rates.push({ maxWeight, postage, code });
    }
  }
  rates.sort((a, b) => a.maxWeight - b.maxWeight);
  return rates;
};

// ============================================================================
// MAIN MESSAGE HANDLER
// ============================================================================

self.onmessage = async (e) => {
  // Check if this is a status request
  if (e.data.type === "check_cache_status") {
    if (categoryIndex.isInitialized) {
      self.postMessage({
        type: "category_cache",
        count: categoryIndex.categories.length,
        isCached: true,
        isNewUpload: false,
      });
    } else {
      self.postMessage({
        type: "category_cache",
        count: 0,
        isCached: false,
        isNewUpload: false,
      });
    }
    return;
  }

  const {
    excelBuffer,
    pdfDataArray,
    categoryBuffer,
    postageBuffer,
    date,
    shippingDiscount,
  } = e.data;

  try {
    // Apply shipping discount equally divided among all PDFs
    if (shippingDiscount && shippingDiscount > 0 && pdfDataArray.length > 0) {
      const discountPerPDF = shippingDiscount / pdfDataArray.length;
      pdfDataArray.forEach((pdf) => {
        pdf.shipping = Math.max(0, pdf.shipping - discountPerPDF);
      });
    }

    // Load Excel
    self.postMessage({ type: "progress", message: "Loading Excel file..." });
    await sleep(0);

    const workbook = XLSX.read(new Uint8Array(excelBuffer));
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    // Load Category Map and BUILD INDEX
    let categoryMap = [];

    if (categoryBuffer) {
      self.postMessage({
        type: "progress",
        message: "Loading & indexing category map...",
      });
      await sleep(0);

      const catWorkbook = XLSX.read(new Uint8Array(categoryBuffer));
      const catSheetName =
        catWorkbook.SheetNames.find(
          (name) =>
            name.toLowerCase().includes("category") ||
            name.toLowerCase().includes("map")
        ) || catWorkbook.SheetNames[0];
      const catWorksheet = catWorkbook.Sheets[catSheetName];
      categoryMap = XLSX.utils.sheet_to_json(catWorksheet);

      categoryIndex.build(categoryMap);
    } else if (!categoryIndex.isInitialized) {
      const catSheetName = workbook.SheetNames.find(
        (name) =>
          name.toLowerCase().includes("category") ||
          name.toLowerCase().includes("map")
      );
      if (catSheetName) {
        const catWorksheet = workbook.Sheets[catSheetName];
        categoryMap = XLSX.utils.sheet_to_json(catWorksheet);

        categoryIndex.build(categoryMap);
      }
    } else {
      self.postMessage({
        type: "category_cache",
        count: categoryIndex.categories.length,
        isCached: true,
        isNewUpload: false,
      });
    }

    // Load Postage Rates
    let postageRates = null;

    if (postageBuffer) {
      self.postMessage({
        type: "progress",
        message: "Loading postage rates...",
      });
      await sleep(0);

      const postageWorkbook = XLSX.read(new Uint8Array(postageBuffer));
      const postageSheetName = postageWorkbook.SheetNames[0];
      const postageWorksheet = postageWorkbook.Sheets[postageSheetName];
      const postageData = XLSX.utils.sheet_to_json(postageWorksheet, {
        header: 1,
      });

      postageRates = parsePostageRates(postageData);
    } else {
      const postageSheetName = workbook.SheetNames.find((name) =>
        name.toLowerCase().includes("postage")
      );
      if (postageSheetName) {
        const postageWorksheet = workbook.Sheets[postageSheetName];
        const postageData = XLSX.utils.sheet_to_json(postageWorksheet, {
          header: 1,
        });
        postageRates = parsePostageRates(postageData);
      }
    }

    // Extract SKUs from Column C (index 2) - will be Column E after shift
    self.postMessage({
      type: "progress",
      message: "Extracting SKUs from Column C...",
    });
    await sleep(0);

    const dataRows = data.slice(1);
    const manifestSKUs = [];

    for (let idx = 0; idx < dataRows.length; idx++) {
      const row = dataRows[idx];
      const sku = (row[2] && String(row[2]).trim()) || "";

      if (sku) {
        manifestSKUs.push({ sku, rowIndex: idx });
      }

      if (idx % 50 === 0) await sleep(0);
    }

    // Search PDFs for SKUs
    self.postMessage({
      type: "progress",
      message: "Searching PDFs for SKU prices...",
    });
    await sleep(0);

    const items = [];

    for (let i = 0; i < manifestSKUs.length; i++) {
      const { sku } = manifestSKUs[i];
      const skuRegex = new RegExp(`\\b${sku.replace(/[-]/g, "[-]?")}\\b`, "gi");
      let found = false;

      for (let pdfData of pdfDataArray) {
        const match = skuRegex.exec(pdfData.text);

        if (match) {
          const skuIndex = match.index;
          const textAfterSKU = pdfData.text.substring(skuIndex, skuIndex + 200);
          const priceMatches = textAfterSKU.match(/£\s*(\d+\.?\d*)/g);

          if (priceMatches && priceMatches.length > 0) {
            const firstPrice = parseFloat(priceMatches[0].replace(/£\s*/g, ""));
            items.push({
              sku: sku,
              cost: firstPrice,
              pdfIndex: pdfData.index,
              pdfName: pdfData.name,
              shipping: pdfData.shipping,
              invoiceDate: pdfData.invoiceDate,
              vendorNumber: pdfData.vendorNumber,
            });
            found = true;
            break;
          }
        }
      }

      if (!found) {
        items.push({
          sku: sku,
          cost: 0,
          pdfIndex: -1,
          pdfName: "Not found",
          shipping: 0,
          invoiceDate: "",
          vendorNumber: "",
        });
      }

      if (i % 20 === 0) await sleep(0);
    }

    const totalShipping = pdfDataArray.reduce(
      (sum, pdf) => sum + pdf.shipping,
      0
    );

    self.postMessage({
      type: "extracted",
      data: { items, totalShipping },
    });

    // Calculate proportional costs
    self.postMessage({
      type: "progress",
      message: "Calculating proportional costs...",
    });
    await sleep(0);

    const skuGroups = {};
    manifestSKUs.forEach(({ sku, rowIndex }) => {
      if (!skuGroups[sku]) skuGroups[sku] = [];
      skuGroups[sku].push(rowIndex);
    });

    const proportionalCosts = {};
    for (const sku of Object.keys(skuGroups)) {
      const rowIndices = skuGroups[sku];
      const pdfItem = items.find((item) => item.sku === sku);
      const invoiceCost = pdfItem ? pdfItem.cost : 0;

      if (rowIndices.length > 1) {
        const totalRRP = rowIndices.reduce((sum, idx) => {
          const rrp = parseFloat(dataRows[idx][23]) || 0;
          return sum + rrp;
        }, 0);

        rowIndices.forEach((idx) => {
          const itemRRP = parseFloat(dataRows[idx][23]) || 0;
          const proportionalCost =
            totalRRP > 0 ? (itemRRP / totalRRP) * invoiceCost : 0;
          proportionalCosts[idx] = proportionalCost;
        });
      } else {
        proportionalCosts[rowIndices[0]] = invoiceCost;
      }

      await sleep(0);
    }

    // Allocate shipping using Currency (Weight × Quantity)
    self.postMessage({
      type: "progress",
      message: "Allocating shipping costs...",
    });
    await sleep(0);

    const itemsByPDF = {};

    dataRows.forEach((row, idx) => {
      const manifestSKU = manifestSKUs[idx]?.sku || "";
      const pdfItem = items.find((item) => item.sku === manifestSKU);
      const pdfIndex = pdfItem ? pdfItem.pdfIndex : -1;
      const pdfShipping = pdfItem ? pdfItem.shipping : 0;

      if (!itemsByPDF[pdfIndex]) {
        itemsByPDF[pdfIndex] = {
          items: [],
          shipping: pdfShipping,
          totalCurrency: 0,
        };
      }

      const weight = parseFloat(row[20]) || 0;
      const quantity = parseInt(row[17]) || 1;
      const currency = weight * quantity; // Currency = Weight × Quantity

      itemsByPDF[pdfIndex].items.push({ rowIndex: idx, currency });
      itemsByPDF[pdfIndex].totalCurrency += currency;
    });

    const itemShipping = {};
    for (const pdfIndex of Object.keys(itemsByPDF)) {
      const pdfGroup = itemsByPDF[pdfIndex];

      pdfGroup.items.forEach(({ rowIndex, currency }) => {
        const shipping =
          pdfGroup.totalCurrency > 0
            ? (currency / pdfGroup.totalCurrency) * pdfGroup.shipping
            : 0;
        itemShipping[rowIndex] = shipping;
      });

      await sleep(0);
    }

    // Build Excel workbook
    self.postMessage({
      type: "progress",
      message: "Building Excel workbook...",
    });
    await sleep(0);

    const dateSKU = formatDateToSKU(date);
    const asinToSKU = {};

    let totalRRP = 0;
    let totalCost = 0;
    let totalShippingAlloc = 0;
    let totalVAT = 0;
    let totalTotalCost = 0;

    // First pass: collect all unique item specifics across all rows
    self.postMessage({
      type: "progress",
      message: "Analyzing item specifics...",
    });
    await sleep(0);

    const allSpecificsKeys = new Set();
    const rowSpecifics = []; // Store specifics for each row

    for (let idx = 0; idx < dataRows.length; idx++) {
      const row = dataRows[idx];
      const originalTitle = row[3] || "";
      const description = String(row[4] || "");
      const specifics = extractItemSpecifics(originalTitle, description);

      rowSpecifics.push(specifics);
      Object.keys(specifics).forEach(key => allSpecificsKeys.add(key));

      if (idx % 50 === 0) await sleep(0);
    }

    // Convert to sorted array for consistent column order
    const specificsColumns = Array.from(allSpecificsKeys).sort();

    const processedRows = [];

    // Build header row WITHOUT C: columns (those go in eBay CSV only)
    const headerRow = [
      "Order Date",
      "Vendor",
      ...data[0].slice(0, 24),
      "Cost",
      "Shipping ",
      "VAT",
      "Total Cost",
      "Cost Per one",
      "Postage",
      "Postage code",
      "SKU",
      "Location",
      "SKU Location ",
      "Shorten Name",
      "",
      "",
      "",
      "Action(SiteID=UK|Country=GB|Currency=GBP|Version=1193|CC=UTF-8)",
      "Custom label (SKU)",
      "Category ID",
      "Title",
      "UPC",
      "Price",
      "Quantity",
      "Item photo URL",
      "Condition ID",
      "Description",
      "Format",
      "",
      "",
      "SKU",
      "Title",
      "Description",
      "Tags",
      "MetaKeywords",
      "MetaDescription",
      "MobileDescription",
      "CategoryID",
      "StoreCategory",
      "PrivateListing",
      "UpToQuantity",
      "WarehouseQuantity",
      "InventoryControl",
      "Price",
      "WholesalePrice",
      "BestOffer",
      "BestOfferAccept",
      "BestOfferDecline",
      "C:MPN",
      "C:Brand",
      "C:Size",
      "Condition",
      "CountryCode",
      "Location",
      "PostalCode",
      "PolicyPayment",
      "PolicyShipping",
      "PolicyReturn",
      "PackageType",
      "MeasurementSystem",
      "PackageLength",
      "PackageWidth",
      "PackageDepth",
      "WeightMajor",
      "WeightMinor",
      "Image 1",
      "Image 2",
      "Image 3",
      "Image 4",
      "Image 5",
      "Image 6",
      "ASIN",
      "ConditionNote",
      "OriginalRetailPrice",
      "Model",
      "EAN",
      "3DsellersCSVTemplateVersion",
      "",
      "",
      "SKU",
      "CONDITION",
      "EBAY Title",
      "BRAND",
    ];
    processedRows.push(headerRow);

    for (let idx = 0; idx < dataRows.length; idx++) {
      const row = dataRows[idx];

      const asin = String(row[5] || "").toLowerCase();
      const weight = parseFloat(row[20]) || 0;
      const quantity = parseInt(row[17]) || 1;
      const rrp = parseFloat(row[22]) || 0;
      const brand = row[8] || ""; // Brand from uploaded excel (column index 8)
      const condition = row[18] || "";
      // Note: subcategory (row[10]) is extracted later in eBay CSV generation

      const manifestSKU = manifestSKUs[idx]?.sku || "";
      const pdfItem = items.find((item) => item.sku === manifestSKU);

      const invoiceDate = pdfItem ? pdfItem.invoiceDate : "";
      const vendorNumber = pdfItem ? pdfItem.vendorNumber : "";

      const cost = round(proportionalCosts[idx] || 0);
      const shipping = round(itemShipping[idx] || 0);

      const vat = round((cost + shipping) * 0.2);
      const totalCostCalc = round(cost + shipping + vat);
      const costPerOne = round(
        quantity > 0 ? totalCostCalc / quantity : totalCostCalc
      );

      // Update SKU to include rounded up cost per one in format: dateSKU/costPerOne/
      const roundedUpCostPerOne = Math.ceil(costPerOne);

      if (asin && !asinToSKU[asin]) {
        asinToSKU[asin] = `${dateSKU}/${roundedUpCostPerOne}/`;
      }
      const sku = asinToSKU[asin] || `${dateSKU}/${roundedUpCostPerOne}/`;
      const postageInfo = getPostageInfo(weight, postageRates);

      const originalTitle = row[3] || "";
      const shortenedTitle = shortenTitle(originalTitle);
      const roundedRRP = Math.ceil(rrp);
      const rrpText = `RRP £${roundedRRP}`;
      const fullTitle = `${shortenedTitle} ${rrpText}`;

      const categoryId = matchCategory(originalTitle);

      const skuLocation = `/${sku}/${postageInfo.code}`;
      const price = round(rrp * 0.85 + 0.15 + postageInfo.postage);
      const bestOffer = round(price * 0.93);

      const imageURL =
        [row[11], row[12], row[13], row[14], row[15], row[16]]
          .filter((img) => img && img !== "N/A")
          .join("|") || "";

      const tags = generateTags(fullTitle);
      const description = String(row[4] || "");
      const metaDesc = `${tags}. ${description.substring(0, 150)}...`;

      totalRRP += rrp;
      totalCost += cost;
      totalShippingAlloc += shipping;
      totalVAT += vat;
      totalTotalCost += totalCostCalc;

      // Build row with Currency replaced (Column V = index 21)
      const rowData = [];
      for (let i = 0; i < 24; i++) {
        if (i === 21) {
          // Column V (index 21) = Currency - replace GBP with Weight × Quantity
          const currencyValue = weight * quantity;
          rowData.push(currencyValue);
        } else {
          rowData.push(row[i]);
        }
      }

      // Build the complete row (NO dynamic C: columns in main Excel)
      const newRow = [
        invoiceDate,
        vendorNumber,
        ...rowData,
        cost,
        shipping,
        vat,
        totalCostCalc,
        costPerOne,
        postageInfo.postage,
        postageInfo.code,
        sku,
        "",
        skuLocation,
        shortenedTitle,
        "RRP £",
        roundedRRP,
        rrpText,
        "Draft",
        skuLocation,
        categoryId,
        fullTitle,
        manifestSKU,
        price,
        quantity,
        imageURL,
        condition,
        row[4] || "",
        "FixedPrice",
        "",
        "",
        skuLocation,
        fullTitle,
        row[4] || "",
        dateSKU,
        tags,
        metaDesc,
        metaDesc,
        20685,
        1,
        "",
        quantity,
        quantity,
        "",
        price,
        costPerOne,
        "true",
        bestOffer,
        "",
        "N/A",
        brand,
        "",
        condition,
        "GB",
        "Dartford",
        "DA4 9EW",
        252103073016,
        254956651016,
        "Return accepted Copy",
        "Package/thick envelope",
        "cm",
        45,
        45,
        16,
        Math.ceil(weight),
        "",
        row[11] || "",
        row[12] || "",
        row[13] || "",
        row[14] || "",
        row[15] || "",
        row[16] || "",
        row[5] || "",
        "",
        rrp,
        brand,
        row[6] || "",
        "S3G TEP",
        "",
        "",
        skuLocation,
        condition,
        fullTitle,
        brand,
      ];

      processedRows.push(newRow);

      if (idx % 50 === 0) {
        self.postMessage({
          type: "progress",
          message: `Building Excel: ${idx + 1}/${dataRows.length} rows...`,
        });
        await sleep(0);
      }
    }

    // Create totals row with dynamic size based on header length
    const totalsRow = new Array(headerRow.length).fill("");
    totalsRow[23] = round(
      dataRows.reduce((sum, row) => {
        const weight = parseFloat(row[20]) || 0;
        const quantity = parseInt(row[17]) || 1;
        return sum + weight * quantity;
      }, 0)
    ); // Total Currency at index 23 (Order Date + Vendor + Column V which is index 21)
    totalsRow[22] = round(
      dataRows.reduce((sum, row) => sum + (parseFloat(row[20]) || 0), 0)
    );
    totalsRow[24] = round(totalRRP);
    totalsRow[25] = round(totalRRP);
    totalsRow[26] = round(totalCost);
    totalsRow[27] = round(totalShippingAlloc);
    totalsRow[28] = round(totalVAT);
    totalsRow[29] = round(totalTotalCost);

    processedRows.push(totalsRow);

    // Create Excel with formulas
    self.postMessage({ type: "progress", message: "Adding Excel formulas..." });
    await sleep(0);

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(processedRows);

    for (let i = 0; i < dataRows.length; i++) {
      const excelRow = i + 2;
      const row = dataRows[i];
      const quantity = parseInt(row[17]) || 1;
      const weight = parseFloat(row[20]) || 0;

      const colRRP = getExcelColumn(24);
      const colQuantity = getExcelColumn(19);
      const colWeight = getExcelColumn(22);
      const colCurrency = getExcelColumn(23); // Column V in output (index 21 + 2 for Order Date + Vendor)
      const colCost = getExcelColumn(26);
      const colShipping = getExcelColumn(27);
      const colVAT = getExcelColumn(28);
      const colTotalCost = getExcelColumn(29);
      const colCostPerOne = getExcelColumn(30);
      const colPostage = getExcelColumn(31);
      const colPrice = getExcelColumn(45);
      const colBestOffer = getExcelColumn(69);

      // Set Currency formula: Weight × Quantity (Column V)
      const currencyFormula = `=${colWeight}${excelRow}*${colQuantity}${excelRow}`;
      newWorksheet[`${colCurrency}${excelRow}`] = {
        t: "n",
        f: currencyFormula,
        v: round(weight * quantity),
      };

      newWorksheet[`${colShipping}${excelRow}`] = {
        t: "n",
        v: round(itemShipping[i] || 0),
      };

      const vatFormula = `=ROUND((${colCost}${excelRow}+${colShipping}${excelRow})*0.2,2)`;
      newWorksheet[`${colVAT}${excelRow}`] = {
        t: "n",
        f: vatFormula,
        v: round(
          (parseFloat(processedRows[i + 1][26]) +
            parseFloat(processedRows[i + 1][27])) *
            0.2
        ),
      };

      const totalCostFormula = `=ROUND(${colCost}${excelRow}+${colShipping}${excelRow}+${colVAT}${excelRow},2)`;
      newWorksheet[`${colTotalCost}${excelRow}`] = {
        t: "n",
        f: totalCostFormula,
        v: round(
          parseFloat(processedRows[i + 1][26]) +
            parseFloat(processedRows[i + 1][27]) +
            parseFloat(processedRows[i + 1][28])
        ),
      };

      const costPerOneFormula = `=ROUND(${colTotalCost}${excelRow}/${colQuantity}${excelRow},2)`;
      newWorksheet[`${colCostPerOne}${excelRow}`] = {
        t: "n",
        f: costPerOneFormula,
        v: round(parseFloat(processedRows[i + 1][29]) / quantity),
      };

      const priceFormula = `=ROUND((${colRRP}${excelRow}*0.85)+0.15+${colPostage}${excelRow},2)`;
      newWorksheet[`${colPrice}${excelRow}`] = {
        t: "n",
        f: priceFormula,
        v: round(
          parseFloat(row[22]) * 0.85 +
            0.15 +
            parseFloat(processedRows[i + 1][31])
        ),
      };

      const bestOfferFormula = `=ROUND(${colPrice}${excelRow}*0.93,2)`;
      newWorksheet[`${colBestOffer}${excelRow}`] = {
        t: "n",
        f: bestOfferFormula,
        v: round(parseFloat(processedRows[i + 1][45]) * 0.93),
      };

      const colPrice2 = getExcelColumn(66);
      newWorksheet[`${colPrice2}${excelRow}`] = {
        t: "n",
        f: priceFormula,
        v: round(
          parseFloat(row[22]) * 0.85 +
            0.15 +
            parseFloat(processedRows[i + 1][31])
        ),
      };

      if (i % 100 === 0) await sleep(0);
    }

    const totalsRowNumber = dataRows.length + 2;

    // Currency column total (Column X in output)
    newWorksheet[`${getExcelColumn(23)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,X2:X${dataRows.length + 1})`,
      v: totalsRow[23],
    };

    newWorksheet[`${getExcelColumn(22)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,W2:W${dataRows.length + 1})`,
      v: totalsRow[22],
    };
    newWorksheet[`${getExcelColumn(24)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,Y2:Y${dataRows.length + 1})`,
      v: totalsRow[24],
    };
    newWorksheet[`${getExcelColumn(25)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,Z2:Z${dataRows.length + 1})`,
      v: totalsRow[25],
    };
    newWorksheet[`${getExcelColumn(26)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,AA2:AA${dataRows.length + 1})`,
      v: totalsRow[26],
    };
    newWorksheet[`${getExcelColumn(27)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,AB2:AB${dataRows.length + 1})`,
      v: totalsRow[27],
    };
    newWorksheet[`${getExcelColumn(28)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,AC2:AC${dataRows.length + 1})`,
      v: totalsRow[28],
    };
    newWorksheet[`${getExcelColumn(29)}${totalsRowNumber}`] = {
      t: "n",
      f: `SUBTOTAL(9,AD2:AD${dataRows.length + 1})`,
      v: totalsRow[29],
    };

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, dateSKU);

    // Add Postage Rate Table
    const finalPostageRates = postageRates || postageRateTable;
    const postageTableData = [
      [
        "Weight (Max)",
        "Postage (£)",
        "Royal Mail Basic",
        "Fuel",
        "Vat",
        "Diff",
        "Code",
      ],
    ];

    finalPostageRates.forEach((rate) => {
      const rmBasic = round(rate.postage / 0.74);
      const fuel = round(rmBasic * 1.08);
      const vat = round(fuel * 1.2);
      const diff = round(vat - rate.postage);

      postageTableData.push([
        rate.maxWeight,
        rate.postage,
        rmBasic,
        fuel,
        vat,
        diff,
        rate.code,
      ]);
    });

    const postageSheet = XLSX.utils.aoa_to_sheet(postageTableData);
    XLSX.utils.book_append_sheet(
      newWorkbook,
      postageSheet,
      "Postage Rate Table"
    );

    // Generate Excel blob
    self.postMessage({ type: "progress", message: "Generating Excel file..." });
    await sleep(0);

    const wbout = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });
    const excelBlob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const excelUrl = URL.createObjectURL(excelBlob);
    const excelFilename = `LISTING_${dateSKU}.xlsx`;

    // Create eBay CSV with C: columns
    self.postMessage({ type: "progress", message: "Generating eBay CSV..." });
    await sleep(0);

    const ebayRows = [];
    ebayRows.push(
      "#INFO,Version=0.0.2,Template= eBay-draft-listings-template_GB,,,,,,,"
    );
    ebayRows.push(
      "#INFO Action and Category ID are required fields. 1) Set Action to Draft 2) Please find the category ID for your listings here: https://pages.ebay.com/sellerinformation/news/categorychanges.html,,,,,,,,,,"
    );
    ebayRows.push(
      "#INFO After you've successfully uploaded your draft from the Seller Hub Reports tab, complete your drafts to active listings here: https://www.ebay.co.uk/sh/lst/drafts,,,,,,,,,"
    );
    ebayRows.push("#INFO,,,,,,,,,,");

    // Build eBay CSV header with C: prefix for ALL specifics (like script3.js)
    const ebayBaseHeader = [
      "Action(SiteID=UK|Country=GB|Currency=GBP|Version=1193|CC=UTF-8)",
      "Custom label (SKU)",
      "Category ID",
      "Title",
      "UPC",
      "Price",
      "Quantity",
      "Item photo URL",
      "Condition ID",
      "Description",
      "Format",
    ];

    // Collect ALL unique specifics keys (Brand, Type, and extracted ones)
    const allEbaySpecificsKeys = new Set();
    allEbaySpecificsKeys.add("Brand");
    allEbaySpecificsKeys.add("Type");
    specificsColumns.forEach(key => allEbaySpecificsKeys.add(key));

    // Sort for consistent column order
    const sortedEbaySpecificsKeys = Array.from(allEbaySpecificsKeys).sort();

    // Add C: prefix to ALL specifics columns
    const ebayHeaderSpecifics = sortedEbaySpecificsKeys.map(key => `C:${key}`);
    const ebayHeaderRow = [...ebayBaseHeader, ...ebayHeaderSpecifics];
    ebayRows.push(ebayHeaderRow.join(","));

    // Helper function to escape CSV values
    const escapeCsv = (val) => {
      if (val === null || val === undefined) return "";
      const str = String(val);
      if (str.includes(",") || str.includes('"') || str.includes("\n")) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    };

    processedRows.slice(1, -1).forEach((row, idx) => {
      const action = "Draft";
      const customLabel = row[33] || "";
      const categoryId = row[42] || 47155;
      const title = row[43] || "";
      const upc = String(row[7] || "").toLowerCase();
      const price = round(row[45]) || 0;
      const quantity = row[46] || 1;
      const photoUrl = row[47] || "";
      const conditionId = 1500;
      const description = String(row[49] || "");
      const format = "FixedPrice";

      // Get brand and subcategory from original data
      const dataRow = dataRows[idx];
      const brand = dataRow[8] || "";
      const subcategory = dataRow[10] || "";

      // Get item specifics for this row
      const itemSpecifics = rowSpecifics[idx] || {};

      // Build complete specifics object including Brand and Type
      const allSpecificsForRow = {
        Brand: brand,
        Type: subcategory,
        ...itemSpecifics
      };

      // Map all specifics values in the sorted order
      const dataSpecifics = sortedEbaySpecificsKeys.map((key) => {
        const value = allSpecificsForRow[key] || "";
        return escapeCsv(value);
      });

      const ebayRowData = [
        action,
        customLabel,
        categoryId,
        escapeCsv(title),
        upc,
        price,
        quantity,
        photoUrl,
        conditionId,
        escapeCsv(description),
        format,
        ...dataSpecifics
      ];

      ebayRows.push(ebayRowData.join(","));
    });

    const csvContent = ebayRows.join("\n");
    const csvBlob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const csvUrl = URL.createObjectURL(csvBlob);
    const csvFilename = `ebay_upload_${dateSKU}.csv`;

    const categoryInfo = categoryIndex.isInitialized
      ? ` Categories matched from ${categoryIndex.categories.length} options.`
      : " Using default category 47155.";
    const postageInfo = postageRates
      ? ` Postage rates from uploaded table (${postageRates.length} tiers).`
      : " Using default postage rates.";

    self.postMessage({
      type: "success",
      message: `✅ Success! Downloaded LISTING_${dateSKU}.xlsx and ebay_upload_${dateSKU}.csv with ${dataRows.length} items!${categoryInfo}${postageInfo}`,
      url: excelUrl,
      filename: excelFilename,
      csvUrl: csvUrl,
      csvFilename: csvFilename,
    });
  } catch (error) {
    self.postMessage({
      type: "error",
      error: error.message,
    });
  }
};
