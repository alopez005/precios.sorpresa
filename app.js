// ============================================================
// PARCHE PriceSync v2 — Soporte multi-plataforma (PrestaShop, WooCommerce, genérico)
// ============================================================
// 
// INSTRUCCIONES:
// 1. En tu app.js, reemplazá la función parseTiendanubeProducts() completa
//    con la nueva parseWebProducts() de abajo.
// 2. En startWebSync(), cambiá la línea:
//      let scraped = parseTiendanubeProducts(html, s.name);
//    por:
//      let scraped = parseWebProducts(html, s.name, s.url);
//    (hay 2 ocurrencias de parseTiendanubeProducts en startWebSync, cambiá ambas)
// 3. En el HTML (index.html), cambiá el tooltip que dice:
//      "Funciona con tiendas en Tiendanube"
//    por:
//      "Funciona con Tiendanube, PrestaShop y otras tiendas online"
// ============================================================


// ===== DETECCIÓN DE PLATAFORMA =====
function detectPlatform(html, url) {
  // PrestaShop indicators
  if (
    html.includes('prestashop') ||
    html.includes('PrestaShop') ||
    html.includes('id_product') ||
    html.includes('product-miniature') ||
    html.includes('product_list') ||
    /controller=search/.test(url || '') ||
    /\/contenido\/\d+/.test(url || '') ||
    html.includes('addToCartUrl') ||
    html.includes('id_product_attribute')
  ) {
    return 'prestashop';
  }

  // Tiendanube indicators
  if (
    html.includes('googleItems') ||
    html.includes('js-item-name') ||
    html.includes('data-product-price') ||
    html.includes('LS.productsCount') ||
    html.includes('tiendanube') ||
    html.includes('Tiendanube')
  ) {
    return 'tiendanube';
  }

  // WooCommerce indicators
  if (
    html.includes('woocommerce') ||
    html.includes('wc-product') ||
    html.includes('product_cat')
  ) {
    return 'woocommerce';
  }

  return 'generic';
}


// ===== PARSER PRESTASHOP =====
function parsePrestaShopProducts(html, supplierName) {
  const products = [];

  // Strategy 1: JSON-LD structured data (many PrestaShop themes include this)
  const jsonLdMatches = html.match(/<script[^>]+type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/gi) || [];
  for (const script of jsonLdMatches) {
    try {
      const raw = script.replace(/<\/?script[^>]*>/gi, '').trim();
      const json = JSON.parse(raw);

      // Could be a single Product, an array, or an ItemList with product elements
      let items = [];
      if (Array.isArray(json)) {
        items = json;
      } else if (json['@type'] === 'ItemList' && Array.isArray(json.itemListElement)) {
        items = json.itemListElement.map(el => el.item || el);
      } else if (json['@type'] === 'Product') {
        items = [json];
      } else if (json['@graph']) {
        items = json['@graph'].filter(g => g['@type'] === 'Product');
      }

      for (const item of items) {
        if (item['@type'] !== 'Product' && !item.name) continue;
        const name = item.name;
        let price = 0;
        if (item.offers) {
          const offers = Array.isArray(item.offers) ? item.offers : [item.offers];
          for (const offer of offers) {
            if (offer.price) { price = parseFloat(offer.price) || 0; break; }
            if (offer.lowPrice) { price = parseFloat(offer.lowPrice) || 0; break; }
          }
        }
        if (name && price > 0) {
          products.push({ name: String(name).trim(), price, supplierName });
        }
      }
    } catch (e) { /* skip invalid JSON-LD */ }
  }

  if (products.length > 0) return deduplicateProducts(products);

  // Strategy 2: PrestaShop product-miniature pattern (PS 1.7+)
  // <article class="product-miniature" data-id-product="..." ...>
  //   <h3 class="product-title"><a href="...">Product Name</a></h3>
  //   <span class="price">XX.XXX,XX $</span>
  const miniaturePattern = /<article[^>]*class="[^"]*product-miniature[^"]*"[^>]*>([\s\S]*?)<\/article>/gi;
  let artMatch;
  while ((artMatch = miniaturePattern.exec(html)) !== null) {
    const block = artMatch[1];
    // Extract name from h3 > a, or from h2 > a, or title attribute
    let name = '';
    const titleLinkMatch = block.match(/<h[23][^>]*class="[^"]*product-title[^"]*"[^>]*>\s*<a[^>]*title="([^"]+)"/i)
      || block.match(/<h[23][^>]*class="[^"]*product-title[^"]*"[^>]*>\s*<a[^>]*>([^<]+)<\/a>/i)
      || block.match(/<a[^>]*class="[^"]*product-name[^"]*"[^>]*title="([^"]+)"/i)
      || block.match(/<a[^>]*class="[^"]*product-name[^"]*"[^>]*>([^<]+)<\/a>/i);
    if (titleLinkMatch) name = titleLinkMatch[1].trim();

    // Extract price
    let price = 0;
    const priceMatch = block.match(/<span[^>]*class="[^"]*price[^"]*"[^>]*>([\s\S]*?)<\/span>/i);
    if (priceMatch) {
      price = parsePrice(priceMatch[1].replace(/<[^>]+>/g, '').trim());
    }

    if (name && price > 0) {
      products.push({ name, price, supplierName });
    }
  }

  if (products.length > 0) return deduplicateProducts(products);

  // Strategy 3: PrestaShop older themes — product-container / product_list
  const productBlockPattern = /<div[^>]*class="[^"]*(?:product-container|product_list_item|ajax_block_product)[^"]*"[^>]*>([\s\S]*?)<\/div>\s*(?:<\/div>)?/gi;
  let pbMatch;
  while ((pbMatch = productBlockPattern.exec(html)) !== null) {
    const block = pbMatch[1];
    let name = '';
    const nameMatch = block.match(/<a[^>]*class="[^"]*product-name[^"]*"[^>]*title="([^"]+)"/i)
      || block.match(/<a[^>]*class="[^"]*product-name[^"]*"[^>]*>([^<]+)<\/a>/i)
      || block.match(/<h[2345][^>]*>\s*<a[^>]*>([^<]+)<\/a>/i);
    if (nameMatch) name = nameMatch[1].trim();

    let price = 0;
    const priceMatch = block.match(/(?:content|itemprop)="price"[^>]*content="([\d.]+)"/i)
      || block.match(/<span[^>]*class="[^"]*(?:product-price|price)[^"]*"[^>]*>([\s\S]*?)<\/span>/i);
    if (priceMatch) {
      price = parsePrice(priceMatch[1].replace(/<[^>]+>/g, '').trim());
    }

    if (name && price > 0) {
      products.push({ name, price, supplierName });
    }
  }

  if (products.length > 0) return deduplicateProducts(products);

  // Strategy 4: Generic — find product links with nearby prices
  // PrestaShop URLs typically look like: /category/ID-slug.html
  return parseGenericProducts(html, supplierName);
}


// ===== PARSER WOOCOMMERCE =====
function parseWooCommerceProducts(html, supplierName) {
  const products = [];

  // JSON-LD first (same as PrestaShop)
  const jsonLdMatches = html.match(/<script[^>]+type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/gi) || [];
  for (const script of jsonLdMatches) {
    try {
      const json = JSON.parse(script.replace(/<\/?script[^>]*>/gi, ''));
      let items = [];
      if (json['@type'] === 'Product') items = [json];
      else if (Array.isArray(json)) items = json.filter(i => i['@type'] === 'Product');
      else if (json['@graph']) items = json['@graph'].filter(g => g['@type'] === 'Product');

      for (const item of items) {
        const name = item.name;
        let price = 0;
        if (item.offers) {
          const offers = Array.isArray(item.offers) ? item.offers : [item.offers];
          price = parseFloat(offers[0]?.price || offers[0]?.lowPrice) || 0;
        }
        if (name && price > 0) products.push({ name: String(name).trim(), price, supplierName });
      }
    } catch (e) { }
  }

  if (products.length > 0) return deduplicateProducts(products);

  // WooCommerce product blocks: <li class="product ...">
  const wooPattern = /<li[^>]*class="[^"]*product[^"]*"[^>]*>([\s\S]*?)<\/li>/gi;
  let m;
  while ((m = wooPattern.exec(html)) !== null) {
    const block = m[1];
    const nameMatch = block.match(/<h[23][^>]*class="[^"]*woocommerce-loop-product__title[^"]*"[^>]*>([^<]+)/i)
      || block.match(/<a[^>]*>\s*<h[23][^>]*>([^<]+)/i);
    const priceMatch = block.match(/<span[^>]*class="[^"]*woocommerce-Price-amount[^"]*"[^>]*>([\s\S]*?)<\/span>/i)
      || block.match(/<span[^>]*class="[^"]*price[^"]*"[^>]*>([\s\S]*?)<\/span>/i);

    if (nameMatch && priceMatch) {
      const name = nameMatch[1].trim();
      const price = parsePrice(priceMatch[1].replace(/<[^>]+>/g, '').trim());
      if (name && price > 0) products.push({ name, price, supplierName });
    }
  }

  return deduplicateProducts(products);
}


// ===== PARSER GENÉRICO (para cualquier tienda) =====
function parseGenericProducts(html, supplierName) {
  const products = [];
  const clean = html.replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<nav[\s\S]*?<\/nav>/gi, '')
    .replace(/<header[\s\S]*?<\/header>/gi, '')
    .replace(/<footer[\s\S]*?<\/footer>/gi, '');

  // Try 1: Find product title links (h2/h3 > a with title) near a price
  const titleAnchorPattern = /<h[2345][^>]*>\s*<a[^>]*(?:title="([^"]+)"|>([^<]+)<\/a>)/gi;
  const found = new Map();
  let match;

  while ((match = titleAnchorPattern.exec(clean)) !== null) {
    const name = (match[1] || match[2] || '').trim();
    if (!name || name.length < 3 || found.has(name.toLowerCase())) continue;

    // Look for a price within ~600 chars before or after
    const startIdx = Math.max(0, match.index - 300);
    const endIdx = Math.min(clean.length, match.index + match[0].length + 600);
    const snippet = clean.substring(startIdx, endIdx);

    // Argentine peso format: XX.XXX,XX $ or $XX.XXX,XX or $XX.XXX
    const priceM = snippet.match(/([\d]{1,3}(?:\.[\d]{3})*(?:,\d{2})?)\s*\$/)
      || snippet.match(/\$\s*([\d]{1,3}(?:\.[\d]{3})*(?:,\d{2})?)/)
      || snippet.match(/\$([\d.,]+)/);
    if (priceM) {
      const price = parsePrice(priceM[1]);
      if (price > 0) {
        found.set(name.toLowerCase(), true);
        products.push({ name, price, supplierName });
      }
    }
  }

  // Try 2: itemprop patterns (schema.org microdata)
  if (products.length === 0) {
    const itemPattern = /itemprop="name"[^>]*>([^<]+)<[\s\S]*?itemprop="price"[^>]*content="([\d.]+)"/gi;
    while ((match = itemPattern.exec(clean)) !== null) {
      const name = match[1].trim();
      const price = parseFloat(match[2]) || 0;
      if (name && price > 0 && !found.has(name.toLowerCase())) {
        found.set(name.toLowerCase(), true);
        products.push({ name, price, supplierName });
      }
    }
  }

  return deduplicateProducts(products);
}


// ===== DEDUPLICACIÓN =====
function deduplicateProducts(products) {
  const seen = new Set();
  return products.filter(p => {
    const key = normDesc(p.name);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}


// ===== PARSER TIENDANUBE (original, sin cambios) =====
function parseTiendanubeProducts(html, supplierName) {
  const products = [];

  // Strategy 1: googleItems JS array
  const giStart = html.indexOf('googleItems');
  if (giStart !== -1) {
    const bracketStart = html.indexOf('[', giStart);
    if (bracketStart !== -1) {
      let depth = 0, end = -1;
      for (let i = bracketStart; i < html.length; i++) {
        if (html[i] === '[') depth++;
        else if (html[i] === ']') { depth--; if (depth === 0) { end = i + 1; break; } }
      }
      if (end > bracketStart) {
        try {
          const items = JSON.parse(html.substring(bracketStart, end));
          for (const item of items) {
            const name = item.info?.item_name;
            const price = item.info?.price;
            if (name && price > 0) {
              products.push({ name: String(name).trim(), price: parseFloat(price) || 0, supplierName });
            }
          }
        } catch (e) { }
      }
    }
  }

  // Strategy 2: js-item-name + data-product-price
  if (products.length === 0) {
    const namePattern = /class="[^"]*js-item-name[^"]*"[^>]*>([\s\S]*?)<\/div>/gi;
    const pricePattern = /data-product-price="(\d+)"/gi;
    const names = [], prices = [];
    let nm;
    while ((nm = namePattern.exec(html)) !== null) {
      const name = nm[1].replace(/<[^>]+>/g, '').trim();
      if (name) names.push(name);
    }
    while ((nm = pricePattern.exec(html)) !== null) {
      prices.push(parseInt(nm[1]) / 100);
    }
    const count = Math.min(names.length, prices.length);
    for (let i = 0; i < count; i++) {
      if (names[i] && prices[i] > 0) {
        products.push({ name: names[i], price: prices[i], supplierName });
      }
    }
  }

  // Strategy 3: JSON-LD
  if (products.length === 0) {
    const jsonLdMatches = html.match(/<script[^>]+type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/gi) || [];
    for (const script of jsonLdMatches) {
      try {
        const json = JSON.parse(script.replace(/<\/?script[^>]*>/gi, ''));
        if (json['@type'] === 'Product' || (Array.isArray(json) && json[0]?.['@type'] === 'Product')) {
          const items = Array.isArray(json) ? json : [json];
          for (const item of items) {
            const name = item.name;
            const price = item.offers?.price || item.offers?.[0]?.price;
            if (name && price) {
              products.push({ name: String(name).trim(), price: parseFloat(price) || 0, supplierName });
            }
          }
        }
      } catch (e) { }
    }
  }

  // Strategy 4: data-product-id DOM pattern
  if (products.length === 0) {
    const clean = html.replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<style[\s\S]*?<\/style>/gi, '');
    const itemPattern = /data-product-id="(\d+)"[\s\S]*?class="[^"]*(?:item-name|product-item-name)[^"]*"[^>]*>([^<]+)<[\s\S]*?\$([\d.,]+)/gi;
    let m;
    while ((m = itemPattern.exec(clean)) !== null) {
      const name = m[2].trim();
      const price = parsePrice(m[3]);
      if (name && price > 0) products.push({ name, price, supplierName });
    }

    // Strategy 5: product link anchors
    if (products.length === 0) {
      const prodLinkPattern = /<a[^>]+href="[^"]*\/productos\/[^"]*"[^>]*(?:title="([^"]+)")?[^>]*>([\s\S]*?)<\/a>/gi;
      const found = new Map();
      let match;
      while ((match = prodLinkPattern.exec(clean)) !== null) {
        const title = match[1] || '';
        const inner = match[2] || '';
        let name = (title || inner.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ')).trim();
        name = name.replace(/\s*-\s*comprar online\s*$/i, '').trim();
        if (!name || name.length < 3 || found.has(name.toLowerCase())) continue;
        const startIdx = match.index;
        const snippet = clean.substring(startIdx, startIdx + 800);
        const priceM = snippet.match(/\$([\d.,]+)/);
        if (priceM) {
          const price = parsePrice(priceM[1]);
          if (price > 0) {
            found.set(name.toLowerCase(), true);
            products.push({ name, price, supplierName });
          }
        }
      }
    }
  }

  return deduplicateProducts(products);
}


// ===== ROUTER PRINCIPAL — reemplaza las llamadas a parseTiendanubeProducts =====
function parseWebProducts(html, supplierName, url) {
  const platform = detectPlatform(html, url);
  console.log(`[PriceSync] Plataforma detectada: ${platform} para ${supplierName}`);

  let products = [];

  switch (platform) {
    case 'tiendanube':
      products = parseTiendanubeProducts(html, supplierName);
      break;
    case 'prestashop':
      products = parsePrestaShopProducts(html, supplierName);
      break;
    case 'woocommerce':
      products = parseWooCommerceProducts(html, supplierName);
      break;
    default:
      // Try all parsers as fallback
      products = parseTiendanubeProducts(html, supplierName);
      if (products.length === 0) products = parsePrestaShopProducts(html, supplierName);
      if (products.length === 0) products = parseWooCommerceProducts(html, supplierName);
      if (products.length === 0) products = parseGenericProducts(html, supplierName);
      break;
  }

  // If detected platform parser failed, try generic as last resort
  if (products.length === 0 && platform !== 'generic') {
    console.log(`[PriceSync] Parser ${platform} no encontró productos, probando genérico...`);
    products = parseGenericProducts(html, supplierName);
  }

  return products;
}


// ============================================================
// CAMBIOS EN startWebSync() — buscar y reemplazar estas líneas
// ============================================================
//
// LÍNEA ORIGINAL (aparece 2 veces):
//   let scraped = parseTiendanubeProducts(html, s.name);
//
// REEMPLAZAR POR:
//   let scraped = parseWebProducts(html, s.name, baseUrl);
//
// TAMBIÉN la segunda ocurrencia (dentro del bloque results_only):
//   scraped = parseTiendanubeProducts(roHtml, s.name);
// REEMPLAZAR POR:
//   scraped = parseWebProducts(roHtml, s.name, baseUrl);
//
// Y para la paginación:
//   const pageScraped = parseTiendanubeProducts(pageHtml, s.name);
// REEMPLAZAR POR:
//   const pageScraped = parseWebProducts(pageHtml, s.name, baseUrl);
//
// ============================================================
// CAMBIO EN EL HTML — tooltip del formulario de proveedores
// ============================================================
//
// LÍNEA ORIGINAL:
//   💡 Funciona con tiendas en <strong>Tiendanube</strong>.
//
// REEMPLAZAR POR:
//   💡 Funciona con <strong>Tiendanube</strong>, <strong>PrestaShop</strong> y otras tiendas online.
//
// ============================================================
// CAMBIO EN startWebSync() — paginación solo para Tiendanube
// ============================================================
//
// El bloque de paginación con LS.productsCount es específico de Tiendanube.
// Para PrestaShop, la paginación usa ?page=N en la URL.
// Se agrega detección de paginación PrestaShop.
// ============================================================

// NOTA: Para la paginación de PrestaShop, agregá este bloque
// DESPUÉS del bloque de paginación de Tiendanube (el de LS.productsCount)
// dentro de startWebSync(), justo antes del logLine de totalForSupplier:

/*
      // PrestaShop pagination: look for ?page=N links
      if (scraped.length > 0 && !countMatch) {
        const psPagePattern = /[?&]page=(\d+)/g;
        let maxPage = 1;
        let psMatch;
        while ((psMatch = psPagePattern.exec(html)) !== null) {
          const pg = parseInt(psMatch[1]);
          if (pg > maxPage) maxPage = pg;
        }
        if (maxPage > 1) {
          logLine(`   ℹ PrestaShop: ${maxPage} páginas detectadas — scrapeando...`);
          let emptyStreak = 0;
          for (let page = 2; page <= maxPage; page++) {
            await new Promise(r => setTimeout(r, 1500));
            try {
              const sep = baseUrl.includes('?') ? '&' : '?';
              const pageUrl = `${baseUrl}${sep}page=${page}`;
              const pageHtml = await fetchPageHTML(pageUrl);
              const pageScraped = parseWebProducts(pageHtml, s.name, pageUrl);
              pageScraped.forEach(p => { p.url = pageUrl; });
              allScraped.push(...pageScraped);
              logLine(`   📄 Página ${page}/${maxPage}: ${pageScraped.length} productos`);
              if (pageScraped.length === 0) {
                emptyStreak++;
                if (emptyStreak >= 3) break;
              } else { emptyStreak = 0; }
            } catch (e) {
              logLine(`   ⚠ Página ${page}: ${e.message}`, 'error');
              emptyStreak++;
              if (emptyStreak >= 3) break;
            }
          }
        }
      }
*/
