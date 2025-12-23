/*
 * js/prestashopApiHandler.js
 * VERSIÓN FINAL CON LÓGICA DE EXTRACCIÓN DE NOMBRES REFORZADA
 */
const PrestashopApiHandler = (() => {
    /**
     * --- CONFIGURACIÓN PERMANENTE ---
     * ID del grupo de atributos para "Talla", identificado mediante el modo de depuración.
     * La aplicación ahora depende de este ID, que es robusto y no cambia.
     */
    const SIZE_ATTRIBUTE_GROUP_ID = 5; // ID para "Talla" confirmado.

    const { baseUrl, apiKeyProductReadWrite, apiKeyImageUpload } = window.APP_CONFIG.prestashop;
    const API_OPTIONS_GET = { headers: { 'Authorization': 'Basic ' + btoa(apiKeyProductReadWrite + ':') } };
    const API_OPTIONS_PUT = { method: 'PUT', headers: { 'Authorization': 'Basic ' + btoa(apiKeyProductReadWrite + ':'), 'Content-Type': 'application/xml' }};
    const API_OPTIONS_POST = { method: 'POST', headers: { 'Authorization': 'Basic ' + btoa(apiKeyProductReadWrite + ':'), 'Content-Type': 'application/xml' }};
    const API_OPTIONS_DELETE = { method: 'DELETE', headers: { 'Authorization': 'Basic ' + btoa(apiKeyProductReadWrite + ':') }};
    const PRODUCTS_PER_PAGE = 20;

    const fetchProducts = async (page = 1, searchQuery = '') => {
        const start = (page - 1) * PRODUCTS_PER_PAGE;
        if (searchQuery) {
            const searchUrl = `${baseUrl}/search?language=1&output_format=JSON&query=${encodeURIComponent(searchQuery)}`;
            const searchResponse = await fetch(searchUrl, API_OPTIONS_GET);
            if (!searchResponse.ok) throw new Error(`Fallo en la búsqueda: ${searchResponse.status}`);
            const searchData = await searchResponse.json();
            const productIds = searchData.products ? searchData.products.map(p => p.id) : [];
            if (productIds.length === 0) return { products: [], totalProducts: 0 };
            const idsForPage = productIds.slice(start, start + PRODUCTS_PER_PAGE);
            if (idsForPage.length === 0) return { products: [], totalProducts: productIds.length };
            const productsUrl = `${baseUrl}/products?display=full&output_format=JSON&filter[id]=[${idsForPage.join('|')}]`;
            const productsResponse = await fetch(productsUrl, API_OPTIONS_GET);
            if (!productsResponse.ok) throw new Error(`Fallo al obtener detalles: ${productsResponse.status}`);
            const productsData = await productsResponse.json();
            return { products: productsData.products || [], totalProducts: productIds.length };
        } else {
            const url = `${baseUrl}/products?display=full&sort=[id_DESC]&limit=${start},${PRODUCTS_PER_PAGE}&output_format=JSON`;
            const response = await fetch(url, API_OPTIONS_GET);
            if (!response.ok) throw new Error(`Fallo en fetchProducts: ${response.status} ${response.statusText}`);
            const totalProductsHeader = response.headers.get('PS-WS-HTTP-ALL-ITEMS-COUNT') || '0';
            const totalProducts = parseInt(totalProductsHeader, 10);
            const data = await response.json();
            return { products: data.products || [], totalProducts: isNaN(totalProducts) ? 0 : totalProducts };
        }
    };

    const fetchAuxData = async () => {
        try {
            const categoriesUrl = `${baseUrl}/categories?display=full&output_format=JSON`;
            const manufacturersUrl = `${baseUrl}/manufacturers?display=full&output_format=JSON`;

            const [catRes, manRes] = await Promise.all([
                fetch(categoriesUrl, API_OPTIONS_GET),
                fetch(manufacturersUrl, API_OPTIONS_GET)
            ]);

            if (!catRes.ok) throw new Error(`Error al cargar CATEGORÍAS: ${catRes.statusText}`);
            if (!manRes.ok) throw new Error(`Error al cargar FABRICANTES: ${manRes.statusText}`);

            const catData = await catRes.json();
            const manData = await manRes.json();

            if (!catData || !Array.isArray(catData.categories)) throw new Error("La respuesta de la API de categorías no tiene el formato esperado.");
            if (!manData || !Array.isArray(manData.manufacturers)) throw new Error("La respuesta de la API de fabricantes no tiene el formato esperado.");

            // --- FUNCIÓN REFORZADA ---
            // Ahora maneja tanto el formato de array multi-idioma como un valor de texto simple.
            const getLocalizedNameFromList = (list) => {
                if (!list) return 'Sin nombre';
                if (Array.isArray(list)) {
                    const nameNode = list.find(n => n.id_language === "1") || list[0];
                    return nameNode?.value || 'Sin nombre';
                }
                // Si no es un array, devuelve el valor directamente. Esto soluciona la inconsistencia de la API.
                return String(list);
            };

            const categoriesMap = catData.categories.reduce((acc, cat) => {
                acc[cat.id] = getLocalizedNameFromList(cat.name);
                return acc;
            }, {});
            const manufacturersMap = manData.manufacturers.reduce((acc, man) => {
                acc[man.id] = man.name || 'Marca sin nombre';
                return acc;
            }, {});

            const sizeAttributeGroupId = SIZE_ATTRIBUTE_GROUP_ID;

            if (!sizeAttributeGroupId) {
                console.error("Error Crítico de Configuración: La constante 'SIZE_ATTRIBUTE_GROUP_ID' no está definida.");
                return { categoriesMap, manufacturersMap, sizeAttributeGroupId: null, allSizeOptionValues: [] };
            }

            const sizeValuesUrl = `${baseUrl}/product_option_values?display=full&output_format=JSON&filter[id_attribute_group]=[${sizeAttributeGroupId}]`;
            const sizeValuesRes = await fetch(sizeValuesUrl, API_OPTIONS_GET);
            if (!sizeValuesRes.ok) throw new Error(`No se pudieron cargar los valores para el grupo de atributos con ID ${sizeAttributeGroupId}.`);
            const sizeValuesData = await sizeValuesRes.json();

            const allSizeOptionValues = (sizeValuesData.product_option_values || []).map(val => ({
                id: val.id,
                value: getLocalizedNameFromList(val.name) // Se usa la nueva función robusta
            }));

            return { categoriesMap, manufacturersMap, sizeAttributeGroupId, allSizeOptionValues };

        } catch (error) {
            console.error("[API Handler] Ocurrió un error dentro de fetchAuxData:", error);
            throw error;
        }
    };

    const fetchProductDetails = async (productId) => {
        const productUrl = `${baseUrl}/products/${productId}?output_format=JSON`;
        const productResponse = await fetch(productUrl, API_OPTIONS_GET);
        if (!productResponse.ok) throw new Error('No se pudo obtener el producto base para los detalles.');
        const productData = (await productResponse.json()).product;

        let combinationsData = {};
        let optionValuesData = {};
        let stockData = {};

        const hasCombinations = productData.associations.combinations && productData.associations.combinations.length > 0;

        // Obtener stock siempre (tanto para productos con combinaciones como sin ellas)
        const stockUrl = `${baseUrl}/stock_availables?display=full&output_format=JSON&filter[id_product]=[${productId}]`;

        const stockRes = await fetch(stockUrl, API_OPTIONS_GET);
        if (stockRes.ok) {
            const stockJson = await stockRes.json();

            (stockJson.stock_availables || []).forEach(s => {
                // Para productos con combinaciones: id_product_attribute !== '0'
                // Para productos sin combinaciones: id_product_attribute === '0'
                stockData[`p${s.id_product}-a${s.id_product_attribute}`] = s;
            });
        }

        if (hasCombinations) {
            const comboIds = productData.associations.combinations.map(c => c.id);
            const combosRes = await fetch(`${baseUrl}/combinations?display=full&output_format=JSON&filter[id]=[${comboIds.join('|')}]`, API_OPTIONS_GET);

            if (!combosRes.ok) throw new Error("Fallo al obtener combinaciones.");

            const combosJson = await combosRes.json();

            (combosJson.combinations || []).forEach(c => { combinationsData[c.id] = c; });

            const optionValueIds = Object.values(combinationsData).flatMap(c => c.associations.product_option_values.map(pov => pov.id));
            if (optionValueIds.length > 0) {
                const ovRes = await fetch(`${baseUrl}/product_option_values?display=full&output_format=JSON&filter[id]=[${[...new Set(optionValueIds)].join('|')}]`, API_OPTIONS_GET);
                if (ovRes.ok) {
                    const ovJson = await ovRes.json();
                    const getLocalizedName = nameField => {
                        if (!nameField) return "N/A";
                        if (Array.isArray(nameField)) { const n = nameField.find(l => l.id_language === "1") || nameField[0]; return n?.value || "N/A"; }
                        return String(nameField);
                    };
                    (ovJson.product_option_values || []).forEach(ov => { optionValuesData[ov.id] = { attributeName: getLocalizedName(ov.name) }; });
                }
            }
        }

        return { product: productData, combinationsData, optionValuesData, stockData };
    };

    const updateProduct = async (productId, productData) => {
        const schemaUrl = `${baseUrl}/products/${productId}`;
        const schemaResponse = await fetch(schemaUrl, { headers: { 'Authorization': 'Basic ' + btoa(apiKeyProductReadWrite + ':') } });
        if (!schemaResponse.ok) throw new Error('No se pudo obtener el schema XML del producto.');
        const xmlString = await schemaResponse.text();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "application/xml");
        
        const langFields = ['name', 'description_short', 'description'];

        Object.keys(productData).forEach(key => {
            if (langFields.includes(key)) {
                // Lógica para campos multi-idioma con CDATA
                const node = xmlDoc.querySelector(`product > ${key} > language[id='1']`);
                if (node) {
                    const cdata = xmlDoc.createCDATASection(productData[key]);
                    node.textContent = ''; 
                    node.appendChild(cdata);
                }
            } else {
                // Lógica para campos simples como 'reference' o 'ean13'
                const node = xmlDoc.querySelector(`product > ${key}`);
                if (node) {
                    node.textContent = productData[key];
                }
            }
        });

        const NODES_TO_REMOVE = ['manufacturer_name', 'position_in_category', 'quantity', 'id_default_combination', 'id_default_image', 'type', 'date_add', 'date_upd', 'associations'];
        NODES_TO_REMOVE.forEach(nodeName => {
            const node = xmlDoc.querySelector(`product > ${nodeName}`);
            if (node) node.remove();
        });
        const serializer = new XMLSerializer();
        const newXmlString = serializer.serializeToString(xmlDoc);
        const updateUrl = `${baseUrl}/products/${productId}`;
        const options = { ...API_OPTIONS_PUT, body: newXmlString };
        const updateResponse = await fetch(updateUrl, options);
        if (!updateResponse.ok) {
            const errorBody = await updateResponse.text();
            throw new Error(`Error al actualizar el producto. Respuesta: ${errorBody}`);
        }
        return { success: true };
    };
    
    const updateProductStatus = async (productId, isActive) => {
        const statusValue = isActive ? '1' : '0';
        const schemaUrl = `${baseUrl}/products/${productId}`;
        const schemaResponse = await fetch(schemaUrl, { headers: { 'Authorization': 'Basic ' + btoa(apiKeyProductReadWrite + ':') } });
        if (!schemaResponse.ok) throw new Error('No se pudo obtener el schema XML del producto.');
        const xmlString = await schemaResponse.text();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "application/xml");
        const activeNode = xmlDoc.querySelector('product > active');
        if (activeNode) activeNode.textContent = statusValue;
        const NODES_TO_REMOVE = ['manufacturer_name', 'position_in_category', 'quantity', 'id_default_combination', 'id_default_image', 'associations', 'type', 'date_add', 'date_upd'];
        NODES_TO_REMOVE.forEach(nodeName => {
            const node = xmlDoc.querySelector(`product > ${nodeName}`);
            if (node) node.remove();
        });
        const serializer = new XMLSerializer();
        const newXmlString = serializer.serializeToString(xmlDoc);
        const updateUrl = `${baseUrl}/products/${productId}`;
        const options = { ...API_OPTIONS_PUT, body: newXmlString };
        const updateResponse = await fetch(updateUrl, options);
        if (!updateResponse.ok) {
            const errorBody = await updateResponse.text();
            throw new Error(`Error al actualizar estado. Respuesta: ${errorBody}`);
        }
        return { success: true };
    };

    const updateStockForCombination = async (productId, combinationId, newQuantity) => {
        const findUrl = `${baseUrl}/stock_availables?output_format=JSON&filter[id_product]=[${productId}]&filter[id_product_attribute]=[${combinationId}]`;
        const findResponse = await fetch(findUrl, API_OPTIONS_GET);
        if (!findResponse.ok) throw new Error('Error buscando el registro de stock para la combinación.');
        const findData = await findResponse.json();
        if (!findData.stock_availables || findData.stock_availables.length === 0) {
            throw new Error(`No se encontró stock para producto ${productId} y combinación ${combinationId}.`);
        }
        
        const stockId = findData.stock_availables[0].id;
        const schemaUrl = `${baseUrl}/stock_availables/${stockId}`;
        const schemaResponse = await fetch(schemaUrl, API_OPTIONS_GET);
        if (!schemaResponse.ok) throw new Error('No se pudo obtener el schema XML del stock.');
        
        const xmlString = await schemaResponse.text();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "application/xml");
        const quantityNode = xmlDoc.querySelector('stock_available > quantity');
        if (quantityNode) {
            quantityNode.textContent = newQuantity;
        }
        
        const serializer = new XMLSerializer();
        const newXmlString = serializer.serializeToString(xmlDoc);

        const updateUrl = `${baseUrl}/stock_availables/${stockId}`;
        const options = { ...API_OPTIONS_PUT, body: newXmlString };
        const updateResponse = await fetch(updateUrl, options);
        if (!updateResponse.ok) {
            const errorBody = await updateResponse.text();
            throw new Error(`Error al actualizar el stock. Respuesta: ${errorBody}`);
        }
        return { success: true };
    };

    const uploadProductImage = async (productId, formData) => {
        const url = `${baseUrl}/images/products/${productId}`;
        const options = {
            method: 'POST',
            headers: { 'Authorization': 'Basic ' + btoa(apiKeyImageUpload + ':') },
            body: formData
        };
        const response = await fetch(url, options);
        if (!response.ok) {
            const errorBody = await response.text();
            throw new Error(`Error al subir imagen. Respuesta: ${errorBody}`);
        }
        return { success: true };
    };

    const updateCoverImage = async (productId, imageId) => {
        const schemaUrl = `${baseUrl}/products/${productId}?output_format=XML`;
        const schemaRes = await fetch(schemaUrl, API_OPTIONS_GET);
        if (!schemaRes.ok) throw new Error('No se pudo obtener el schema XML del producto.');
        
        const xmlString = await schemaRes.text();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "application/xml");
        
        const coverNode = xmlDoc.querySelector('product > id_default_image');
        if (coverNode) {
            coverNode.setAttribute('xlink:href', `${baseUrl}/images/products/${productId}/${imageId}`);
            coverNode.textContent = imageId;
        } else {
             const newCoverNode = xmlDoc.createElement('id_default_image');
            newCoverNode.setAttribute('xlink:href', `${baseUrl}/images/products/${productId}/${imageId}`);
            newCoverNode.textContent = imageId;
            xmlDoc.querySelector('product').appendChild(newCoverNode);
        }
        
        const NODES_TO_REMOVE = ['manufacturer_name', 'position_in_category', 'quantity', 'id_default_combination', 'type', 'date_add', 'date_upd', 'associations'];
        NODES_TO_REMOVE.forEach(nodeName => {
            const node = xmlDoc.querySelector(`product > ${nodeName}`);
            if (node) node.remove();
        });

        const serializer = new XMLSerializer();
        const newXmlString = serializer.serializeToString(xmlDoc);
        
        const updateUrl = `${baseUrl}/products/${productId}`;
        const options = { ...API_OPTIONS_PUT, body: newXmlString };
        const updateRes = await fetch(updateUrl, options);
        
        if (!updateRes.ok) {
            const errorBody = await updateRes.text();
            throw new Error(`Error al actualizar la imagen de portada. Respuesta: ${errorBody}`);
        }
        return { success: true };
    };

    const deleteProductImage = async (productId, imageId) => {
        const url = `${baseUrl}/images/products/${productId}/${imageId}`;
        const response = await fetch(url, API_OPTIONS_DELETE);
        if (!response.ok) {
            const errorBody = await response.text();
            throw new Error(`Error al eliminar la imagen. Respuesta: ${errorBody}`);
        }
        return { success: true };
    };

    const createCombination = async (productId, productOptionValueId, initialStock) => {
        const combinationXml = `
            <prestashop>
                <combination>
                    <id_product>${productId}</id_product>
                    <minimal_quantity>1</minimal_quantity>
                    <associations>
                        <product_option_values>
                            <product_option_value>
                                <id>${productOptionValueId}</id>
                            </product_option_value>
                        </product_option_values>
                    </associations>
                </combination>
            </prestashop>`;
        
        const comboResponse = await fetch(`${baseUrl}/combinations`, { ...API_OPTIONS_POST, body: combinationXml });
        if (!comboResponse.ok) {
             throw new Error(`Error creando la combinación: ${await comboResponse.text()}`);
        }
        const comboXmlText = await comboResponse.text();
        const newCombinationId = new DOMParser().parseFromString(comboXmlText, "application/xml").querySelector('combination > id').textContent;
        
        if (parseInt(initialStock, 10) > 0) {
            await updateStockForCombination(productId, newCombinationId, initialStock);
        }
        
        return { success: true };
    };

    const deleteCombination = async (combinationId) => {
        const url = `${baseUrl}/combinations/${combinationId}`;
        const response = await fetch(url, API_OPTIONS_DELETE);
        if (!response.ok) {
            const errorBody = await response.text();
            throw new Error(`Error al eliminar la combinación. Respuesta: ${errorBody}`);
        }
        return { success: true };
    };

    return {
        fetchProducts,
        fetchAuxData,
        fetchProductDetails,
        updateProductStatus,
        updateProduct,
        updateStockForCombination,
        uploadProductImage,
        updateCoverImage,
        deleteProductImage,
        createCombination,
        deleteCombination,
        getBaseUrl: () => baseUrl,
        getImageApiKey: () => apiKeyImageUpload || apiKeyProductReadWrite,
        PRODUCTS_PER_PAGE
    };
})();