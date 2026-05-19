const CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSkxL-rHfCyrQzP6tgmXzUd9BCtJ4NcSqiwN57k6pM72YzCLQN16CyQW3-yJE2GeloK9_J8DeeFyL5p/pub?output=csv';

let allProductsData = [];

const brandAliases = {
    'elf': 'Elfbar',
    'elfbar': 'Elfbar',
    'lost': 'Lostmary',
    'lostmary': 'Lostmary',
    'geek': 'Geekvape',
    'geekvape': 'Geekvape'
};

// ==========================================================================
// СЛОВАРИ КАРТИНОК ПО КАТЕГОРИЯМ (БРЕНДАМ)
// ==========================================================================
const imagesByBrand = {
    'elfbar': {
        'touch': 'touch_10000.png',
        'triplex': 'triplex_30000.png',
        'sour king': 'sourking_30000.png',
        'sourking': 'sourking_30000.png',
        'sweet king': 'sweetking_30000.png',
        'sweetking': 'sweetking_30000.png',
        'lush king': 'lushking_35000.png',
        'lushking': 'lushking_35000.png',
        'ice king': 'iceking_30000.png',
        'iceking': 'iceking_30000.png',
        'moonnight': 'moonnight_25000.png',
        'bc': 'bc_18000.png',
        'fs': 'fs_18000.png',
        'cigalike': 'cigalike_600.png',
        '23.000': '23000.png',
        '30.000': '30000.png',
        'nic': 'nicking.png'
    },
    'waka': {
        'prd 60000': 'prd_60000.png',
        'waka 2500': '2500.png',
        'light 8000': 'light_8000.png',
        'light 20000': 'light_20000.png',
        'light 25000': 'light_25000.png',
        'waka 20000': '20000.png',
        'slim': 'slim.png',
        'sopro': 'sopro.png',
        'nova': 'nova.png',
        'extra': 'extra.png',
        'waka 35000': '35000.png',
    },
    'lostmary': {
        ' mo 5000': 'mo_5000.png',
        ' cm': 'cm.png',
        ' ultra os': 'ultra_os.png',
        ' cf': 'cf.png',
        ' black': 'black.png',
        ' cd': 'cd.png',
        ' mo 10000': 'mo_10000.png',
        ' os 12000': 'os_12000.png',
        ' mt': 'mt.png',
        ' bm 16000 -новый дизайн': 'bm_new.png',
        ' bm 16000': 'bm.png',
        ' x-link': 'x-link.png',
        ' combo': 'combo.png',
        ' mixer+': 'mixer.png',
        ' os25000': 'os25000.png',
        ' m0 30000+': 'm0.png',
        'mary 30000': '30000.png',
        ' puff ball': 'puffball.png',
    },
    'puffmi': {
        'dura': 'dura.png',
        'pure': 'pure.png',
        'tank': 'tank.png',
        'flora': 'flora.png',
    },
    'plonq': {
        'alpha': 'alpha.png',
        'plus 1500': 'plus_1500.png',
        'plus pro': 'plus_pro.png',
        'max 6000': 'max_6000.png',
        'max smart': 'max_smart.png',
        'roqy m': 'roqy_m.png',
        'max pro': 'max_pro.png',
        'prime': 'prime.png',
        'ultra': 'ultra.png',
        'roqy l': 'roqy_l.png',
    },
    'geekvape': {
        '25000': '25000.png',
        '32000': '32000.png',
        '40000': '40000.png',
        '50000': '50000.png',
        '60000': '60000.png',
    },
    'vozol': {
        'vista': 'vista.png',
        'gear 20000': 'gear_20000.png',
        'star 20000': 'star_20000.png',
        'gear shisha': 'gear_shisha.png',
        'rave 46000': 'rave_46000.png',
        'star 46000': 'star_46000.png',
        'rave 56000': 'rave_56000.png',
    },
};

// ==========================================================================
// ФУНКЦИЯ ДИНАМИЧЕСКОГО РАСЧЕТА СТОИМОСТИ (ЗОЛОТАЯ СЕРЕДИНА + ОКРУГЛЕНИЕ)
// ==========================================================================
function calculateSitePrice(basePrice) {
    if (!basePrice || isNaN(basePrice)) return 0;

    const k = 3;
    const dynamicK = 1 + (k - 1) / (1 + (basePrice / 380));
    let margin = basePrice * (dynamicK - 1);
    margin = Math.max(300, Math.min(700, margin));

    const finalPrice = basePrice + margin;
    return Math.round(finalPrice / 100) * 100;
}

async function fetchProducts() {
    try {
        const response = await fetch(CSV_URL);
        const data = await response.text();
        const workbook = XLSX.read(data, { type: 'string' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(sheet).slice(2);

        const emojiRegex = /[\u{1F600}-\u{1F64F}\u{1F300}-\u{1F5FF}\u{1F680}-\u{1F6FF}\u{1F700}-\u{1F77F}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F900}-\u{1F9FF}\u{1F1E6}-\u{1F1FF}]/u;

        const baseFiltered = rawData.map(item => {
            const rawFullName = String(item['__EMPTY'] || '').trim();
            const rawPrice = String(item['__EMPTY_1'] || '').trim();
            const cleanPrice = parseInt(rawPrice.replace(/[^0-9]/g, '')) || 0;

            const finalSitePrice = calculateSitePrice(cleanPrice);

            const puffMatch = rawFullName.match(/\d{1,2}\.\d{3}|\d{3,5}/);
            let displayPuffs = null;
            let puffValueForFilter = 0;

            if (puffMatch) {
                displayPuffs = puffMatch[0].replace('.', '');
                puffValueForFilter = parseInt(displayPuffs);
            }

            let displayName = rawFullName;
            if (puffMatch) {
                displayName = rawFullName.replace(puffMatch[0], '').replace(/\s+/g, ' ').trim();
            }

            let originalBrand = rawFullName.split(' ')[0];
            let searchKey = originalBrand.toLowerCase().replace(/[^a-zа-я0-9]/gi, '');
            let finalBrand = brandAliases[searchKey] || originalBrand;
            let normalizedBrand = finalBrand.toLowerCase().replace(/[^a-zа-я0-9]/gi, '');

            // ПОДБОР КАРТИНКИ ИЗ СЛОВАРЯ КАТЕГОРИИ
            let productImg = null;
            const lowerFullName = rawFullName.toLowerCase();
            console.log(lowerFullName);
            // Проверяем, есть ли у нас отдельный словарь под этот бренд
            const brandMap = imagesByBrand[normalizedBrand];
            if (brandMap) {
                // Ищем ключ именно в словаре этого бренда
                for (const [key, fileName] of Object.entries(brandMap)) {
                    if (lowerFullName.includes(key.toLowerCase())) {
                        productImg = `images/${normalizedBrand}/${fileName}`;
                        break;
                    }
                }
            }

            return {
                name: displayName,
                puffs: displayPuffs,
                puffValue: puffValueForFilter,
                brand: finalBrand,
                normBrand: normalizedBrand,
                price: finalSitePrice,
                cleanPrice: cleanPrice,
                img: productImg,
                hasEmoji: emojiRegex.test(rawFullName)
            };
        }).filter(p => !p.hasEmoji && p.cleanPrice !== 0 && p.name.length > 2);

        const brandCounts = {};
        baseFiltered.forEach(p => {
            if (p.normBrand) brandCounts[p.normBrand] = (brandCounts[p.normBrand] || 0) + 1;
        });

        const seenNames = new Set();
        allProductsData = [];

        baseFiltered.forEach(p => {
            if (brandCounts[p.normBrand] > 3) {
                const itemKey = (p.name + (p.puffs || '')).toLowerCase().replace(/[^a-zа-я0-9]/gi, '');
                if (!seenNames.has(itemKey)) {
                    allProductsData.push(p);
                    seenNames.add(itemKey);
                }
            }
        });

        populateBrandFilter(allProductsData);
        renderProducts(allProductsData);
    } catch (e) { console.error(e); }
}

function renderProducts(products) {
    const grid = document.getElementById('productGrid');
    grid.innerHTML = '';

    if (products.length === 0) {
        grid.innerHTML = '<div class="loader">Товары не найдены</div>';
        return;
    }

    products.forEach(p => {
        const card = document.createElement('div');
        card.className = 'product-card';

        const imgTag = p.img
            ? `<img src="${p.img}" alt="${p.name}">`
            : `<div class="no-image">📸</div>`;

        card.innerHTML = `
            <div class="product-img-container">
                ${imgTag}
            </div>
            <div class="product-info">
                <h3 class="product-title">${p.name}</h3>
                <div class="product-details">
                    <span class="puffs-count">${p.puffs ? p.puffs + ' затяжек' : ''}</span>
                    <span class="price">${p.price} ₽</span>
                </div>
            </div>
        `;
        grid.appendChild(card);
    });
}

function populateBrandFilter(products) {
    const brandSelect = document.getElementById('brandFilter');
    brandSelect.innerHTML = '<option value="">Все бренды</option>';
    const uniqueBrandsMap = {};
    products.forEach(p => {
        if (p.normBrand && !uniqueBrandsMap[p.normBrand]) uniqueBrandsMap[p.normBrand] = p.brand;
    });
    Object.keys(uniqueBrandsMap).sort().forEach(norm => {
        const opt = document.createElement('option');
        opt.value = norm;
        opt.textContent = uniqueBrandsMap[norm];
        brandSelect.appendChild(opt);
    });
}

function filterProducts() {
    const searchQuery = document.getElementById('searchInput').value.toLowerCase().trim();
    const selectedNormBrand = document.getElementById('brandFilter').value;
    const minPuffs = parseInt(document.getElementById('puffRange').value) || 0;
    const sortValue = document.getElementById('sortFilter').value;

    document.getElementById('puffValue').textContent = minPuffs > 0 ? minPuffs + '+' : 'Все';

    let filtered = allProductsData.filter(p => {
        const matchesSearch = p.name.toLowerCase().includes(searchQuery);
        const matchesBrand = selectedNormBrand === "" || p.normBrand === selectedNormBrand;
        const matchesPuffs = minPuffs === 0 || (p.puffValue >= minPuffs);
        return matchesSearch && matchesBrand && matchesPuffs;
    });

    if (sortValue === 'priceAsc') {
        filtered.sort((a, b) => a.price - b.price);
    } else if (sortValue === 'priceDesc') {
        filtered.sort((a, b) => b.price - a.price);
    } else if (sortValue === 'puffsAsc') {
        filtered.sort((a, b) => a.puffValue - b.puffValue);
    } else if (sortValue === 'puffsDesc') {
        filtered.sort((a, b) => b.puffValue - a.puffValue);
    }

    renderProducts(filtered);
}

window.onload = fetchProducts;