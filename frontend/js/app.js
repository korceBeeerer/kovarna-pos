// Konfigurace
const API_URL = window.location.origin;

// Stav aplikace
let products = [];
let cart = [];
let stornoMode = false; // true = položky v košíku se z Excelu odečtou

// DOM elementy
const productsGrid = document.getElementById('productsGrid');
const cartItems = document.getElementById('cartItems');
const totalPrice = document.getElementById('totalPrice');
const clearBtn = document.getElementById('clearBtn');
const saveBtn = document.getElementById('saveBtn');
const customOrderBtn = document.getElementById('customOrderBtn');
const backupBtn = document.getElementById('backupBtn');
const eurRateInput = document.getElementById('eurRate');
const customOrderModal = document.getElementById('customOrderModal');
const closeModal = document.getElementById('closeModal');
const customOrderForm = document.getElementById('customOrderForm');
const serverStatus = document.getElementById('serverStatus');
const statusText = document.getElementById('statusText');
const notification = document.getElementById('notification');
const stornoModeBtn = document.getElementById('stornoModeBtn');
const cartPanel = document.getElementById('cartPanel');
const cartTitle = document.getElementById('cartTitle');
const settingsBtn = document.getElementById('settingsBtn');
const settingsModal = document.getElementById('settingsModal');
const closeSettings = document.getElementById('closeSettings');
const settingsForm = document.getElementById('settingsForm');
const excelPathInput = document.getElementById('excelPath');
const excelPathStatus = document.getElementById('excelPathStatus');

// Funkce pro získání EUR kurzu
function getEurRate() {
    const rate = parseFloat(eurRateInput.value) || 25;
    return rate;
}

// Funkce pro převod CZK na EUR
function convertToEur(czk) {
    return czk / getEurRate();
}

// Inicializace
document.addEventListener('DOMContentLoaded', () => {
    loadProducts();
    loadConfig();
    checkServerConnection();
    setupEventListeners();
    renderCart();
});

// Načtení nastavení - pokud Excel soubor neexistuje, rovnou otevři nastavení
async function loadConfig() {
    try {
        const response = await fetch(`${API_URL}/api/config`);
        const data = await response.json();
        if (data.success) {
            excelPathInput.value = data.excel_file || '';
            updateExcelPathStatus(data.exists);
            if (!data.exists) {
                openSettingsModal();
                showNotification('Excel soubor nenalezen - nastavte cestu k souboru', 'error');
            }
        }
    } catch (error) {
        console.error('Chyba při načítání nastavení:', error);
    }
}

function updateExcelPathStatus(exists) {
    excelPathStatus.textContent = exists ? '✓ Soubor nalezen' : '✗ Soubor nenalezen';
    excelPathStatus.className = `settings-status ${exists ? 'ok' : 'error'}`;
}

function openSettingsModal() {
    settingsModal.classList.add('active');
}

function closeSettingsModal() {
    settingsModal.classList.remove('active');
}

// Uložení nové cesty k Excelu
async function saveSettings(event) {
    event.preventDefault();

    try {
        const response = await fetch(`${API_URL}/api/config`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ excel_file: excelPathInput.value })
        });

        const result = await response.json();

        if (response.ok && result.success) {
            excelPathInput.value = result.excel_file;
            updateExcelPathStatus(result.exists);
            closeSettingsModal();
            showNotification('Cesta k Excelu uložena, načítám produkty', 'success');
            loadProducts();
        } else {
            showNotification(result.error || 'Neznámá chyba', 'error');
        }
    } catch (error) {
        showNotification('Chyba připojení k serveru: ' + error.message, 'error');
    }
}

// Kontrola připojení k serveru
async function checkServerConnection() {
    try {
        const response = await fetch(`${API_URL}/api/health`);
        if (response.ok) {
            updateServerStatus(true);
        } else {
            updateServerStatus(false);
        }
    } catch (error) {
        updateServerStatus(false);
    }
}

function updateServerStatus(connected) {
    const indicator = serverStatus.querySelector('.status-indicator');
    if (connected) {
        indicator.classList.add('connected');
        indicator.classList.remove('error');
        statusText.textContent = 'Připojeno k serveru';
    } else {
        indicator.classList.remove('connected');
        indicator.classList.add('error');
        statusText.textContent = 'Chyba připojení';
    }
}

// Načtení produktů z Excelu (přes API)
async function loadProducts() {
    try {
        const response = await fetch(`${API_URL}/api/products`);
        if (response.ok) {
            const data = await response.json();
            if (data.success && data.products) {
                products = data.products;
                renderProducts();
                console.log(`Načteno ${data.products.length} produktů z Excelu`);
            } else {
                showNotification('Chyba při načítání produktů: ' + (data.error || 'Neznámá chyba'), 'error');
                loadDefaultProducts();
            }
        } else {
            showNotification('Server není dostupný, používám výchozí produkty', 'error');
            loadDefaultProducts();
        }
    } catch (error) {
        console.error('Chyba při načítání produktů:', error);
        showNotification('Chyba připojení, používám výchozí produkty', 'error');
        loadDefaultProducts();
    }
}

// Záložní výchozí produkty (pokud Excel není dostupný)
function loadDefaultProducts() {
    products = [
        { id: 1, name: 'Podkova', price: 150 },
        { id: 2, name: 'Klíč malý', price: 200 },
        { id: 3, name: 'Klíč velký', price: 350 },
        { id: 4, name: 'Nůž', price: 500 },
        { id: 5, name: 'Mečík', price: 1200 },
        { id: 6, name: 'Sekera', price: 800 }
    ];
    renderProducts();
}

// Vykreslení produktů - textová okýnka bez obrázků
function renderProducts() {
    productsGrid.innerHTML = '';

    products.forEach(product => {
        const card = document.createElement('div');
        card.className = 'product-card';
        card.onclick = () => addToCart(product);

        card.innerHTML = `
            <div class="product-name">${product.name}</div>
            <div class="product-price">${product.price} Kč</div>
        `;

        productsGrid.appendChild(card);
    });
}

// Přidání do košíku
function addToCart(product) {
    const existingItem = cart.find(item => item.id === product.id);

    if (existingItem) {
        existingItem.quantity++;
    } else {
        cart.push({
            id: product.id,
            name: product.name,
            price: product.price,
            quantity: 1
        });
    }

    renderCart();
    showNotification(stornoMode ? 'Přidáno ke stornu' : 'Přidáno do košíku', 'success');
}

// Vykreslení košíku
function renderCart() {
    if (cart.length === 0) {
        cartItems.innerHTML = stornoMode
            ? '<p class="empty-cart">Vyberte produkty, které chcete stornovat</p>'
            : '<p class="empty-cart">Košík je prázdný</p>';
        totalPrice.innerHTML = '0 Kč<br><span class="price-eur">0 EUR</span>';
        saveBtn.disabled = true;
        return;
    }

    saveBtn.disabled = false;
    cartItems.innerHTML = '';

    cart.forEach(item => {
        const cartItem = document.createElement('div');
        cartItem.className = 'cart-item';

        const itemTotal = item.price * item.quantity;
        const itemTotalEur = convertToEur(itemTotal);

        cartItem.innerHTML = `
            <div class="cart-item-info">
                <div class="cart-item-name">${item.name}</div>
                <div class="cart-item-price">
                    ${item.price} Kč / ks
                    <span class="price-eur-small">(${convertToEur(item.price).toFixed(2)} EUR)</span>
                </div>
                <div class="cart-item-total">
                    ${itemTotal} Kč
                    <span class="price-eur-small">(${itemTotalEur.toFixed(2)} EUR)</span>
                </div>
            </div>
            <div class="cart-item-controls">
                <button class="quantity-btn" onclick="decreaseQuantity(${item.id})">-</button>
                <span class="quantity">${item.quantity}</span>
                <button class="quantity-btn" onclick="increaseQuantity(${item.id})">+</button>
            </div>
        `;

        cartItems.appendChild(cartItem);
    });

    updateTotal();
}

// Zvýšení množství
function increaseQuantity(productId) {
    const item = cart.find(item => item.id === productId);
    if (item) {
        item.quantity++;
        renderCart();
    }
}

// Snížení množství
function decreaseQuantity(productId) {
    const item = cart.find(item => item.id === productId);
    if (item) {
        item.quantity--;
        if (item.quantity <= 0) {
            cart = cart.filter(i => i.id !== productId);
        }
        renderCart();
    }
}

// Aktualizace celkové ceny
function updateTotal() {
    const total = cart.reduce((sum, item) => sum + (item.price * item.quantity), 0);
    const totalEur = convertToEur(total);
    const sign = stornoMode ? '−' : '';
    totalPrice.innerHTML = `${sign}${total} Kč<br><span class="price-eur">${sign}${totalEur.toFixed(2)} EUR</span>`;
}

// Vymazání košíku
function clearCart() {
    if (cart.length === 0) return;

    if (confirm('Opravdu chcete vymazat všechny položky z košíku?')) {
        cart = [];
        renderCart();
        showNotification('Košík byl vyprázdněn', 'info');
    }
}

// Přepnutí mezi prodejem a stornem
function toggleStornoMode() {
    if (cart.length > 0 && !confirm('Přepnutím režimu se vyprázdní košík. Pokračovat?')) {
        return;
    }

    stornoMode = !stornoMode;
    cart = [];

    cartPanel.classList.toggle('storno-mode', stornoMode);
    stornoModeBtn.classList.toggle('active', stornoMode);
    stornoModeBtn.textContent = stornoMode ? 'Zpět na prodej' : 'Storno';
    cartTitle.textContent = stornoMode ? 'Storno' : 'Košík';
    saveBtn.textContent = getSaveButtonText();

    renderCart();
    showNotification(stornoMode ? 'Režim storna zapnut' : 'Režim prodeje', 'info');
}

function getSaveButtonText() {
    return stornoMode ? 'Stornovat položky' : 'Uložit objednávku';
}

// Uložení objednávky (nebo storna)
async function saveOrder() {
    if (cart.length === 0) {
        showNotification('Košík je prázdný', 'error');
        return;
    }

    if (stornoMode) {
        const summary = cart.map(item => `${item.quantity}× ${item.name}`).join('\n');
        if (!confirm(`Opravdu stornovat tyto položky z dnešního dne?\n\n${summary}`)) {
            return;
        }
    }

    const total = cart.reduce((sum, item) => sum + (item.price * item.quantity), 0);

    const orderData = {
        items: cart,
        total: total
    };

    try {
        saveBtn.disabled = true;
        saveBtn.textContent = 'Ukládání...';

        const endpoint = stornoMode ? 'storno' : 'order';
        const response = await fetch(`${API_URL}/api/${endpoint}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(orderData)
        });

        const result = await response.json();

        if (response.ok && result.success) {
            showNotification(stornoMode ? 'Položky byly stornovány!' : 'Objednávka byla uložena!', 'success');
            cart = [];
            renderCart();
        } else {
            showNotification('Chyba při ukládání: ' + (result.error || 'Neznámá chyba'), 'error');
        }
    } catch (error) {
        showNotification('Chyba připojení k serveru: ' + error.message, 'error');
        updateServerStatus(false);
    } finally {
        saveBtn.disabled = cart.length === 0;
        saveBtn.textContent = getSaveButtonText();
    }
}

// Přidání zakázky
function openCustomOrderModal() {
    customOrderModal.classList.add('active');
}

function closeCustomOrderModal() {
    customOrderModal.classList.remove('active');
    customOrderForm.reset();
}

function addCustomOrder(event) {
    event.preventDefault();

    const name = document.getElementById('customName').value;
    const price = parseInt(document.getElementById('customPrice').value);
    const quantity = parseInt(document.getElementById('customQuantity').value);

    // Vytvoř speciální položku přímo do košíku
    const customItem = {
        id: Date.now(),
        name: name,
        price: price,
        quantity: quantity
    };

    cart.push(customItem);
    renderCart();
    closeCustomOrderModal();
    showNotification(`Zakázka "${name}" přidána ${stornoMode ? 'ke stornu' : 'do košíku'}`, 'success');
}

// Zobrazení notifikace
function showNotification(message, type = 'info') {
    notification.textContent = message;
    notification.className = `notification ${type} show`;

    setTimeout(() => {
        notification.classList.remove('show');
    }, 3000);
}

// Vytvoření zálohy Excelu
async function createBackup() {
    try {
        backupBtn.disabled = true;
        backupBtn.innerHTML = '<span class="backup-icon">⏳</span><span class="backup-text">Zálohování...</span>';

        const response = await fetch(`${API_URL}/api/backup`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });

        const result = await response.json();

        if (response.ok && result.success) {
            showNotification('Záloha úspěšně vytvořena: ' + result.backup_filename, 'success');
        } else {
            showNotification('Chyba při zálohování: ' + (result.message || result.error || 'Neznámá chyba'), 'error');
        }
    } catch (error) {
        showNotification('Chyba připojení k serveru: ' + error.message, 'error');
    } finally {
        backupBtn.disabled = false;
        backupBtn.innerHTML = '<span class="backup-icon">💾</span><span class="backup-text">Zálohovat</span>';
    }
}

// Event listeners
function setupEventListeners() {
    clearBtn.addEventListener('click', clearCart);
    saveBtn.addEventListener('click', saveOrder);
    backupBtn.addEventListener('click', createBackup);
    customOrderBtn.addEventListener('click', openCustomOrderModal);
    closeModal.addEventListener('click', closeCustomOrderModal);
    customOrderForm.addEventListener('submit', addCustomOrder);
    stornoModeBtn.addEventListener('click', toggleStornoMode);
    settingsBtn.addEventListener('click', openSettingsModal);
    closeSettings.addEventListener('click', closeSettingsModal);
    settingsForm.addEventListener('submit', saveSettings);

    // Zavření nastavení kliknutím mimo okno
    settingsModal.addEventListener('click', (e) => {
        if (e.target === settingsModal) {
            closeSettingsModal();
        }
    });

    // Přepočítání cen při změně EUR kurzu
    eurRateInput.addEventListener('input', () => {
        renderCart(); // Přepočítá košík s novým kurzem
    });

    // Zavření modalu kliknutím mimo něj
    customOrderModal.addEventListener('click', (e) => {
        if (e.target === customOrderModal) {
            closeCustomOrderModal();
        }
    });

    // Pravidelná kontrola připojení
    setInterval(checkServerConnection, 30000); // Každých 30 sekund
}

// Globální funkce pro HTML onclick
window.addToCart = addToCart;
window.increaseQuantity = increaseQuantity;
window.decreaseQuantity = decreaseQuantity;
