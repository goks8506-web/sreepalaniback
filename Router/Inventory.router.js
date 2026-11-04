const express = require('express');
const router = express.Router();
const { addProduct, getProducts, addProductType, getProductTypes, updateProduct, deleteProduct, toggleProductStatus, toggleFastRunning, deleteProductType, toggleFree, bulkUploadProductImages, bulkUpdateDiscount } = require('../Controller/Inventory.controller');
const axios = require('axios');
const { Pool } = require('pg');

const pool = new Pool({
  user: process.env.PGUSER,
  password: process.env.PGPASSWORD,
  host: process.env.PGHOST,
  port: process.env.PGPORT,
  database: process.env.PGDATABASE,
});
const multer = require('multer');
const { storage } = require('../Config/cloudinary');

const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB limit
});

router.post('/products', upload.array('images'), addProduct);
router.get('/products', getProducts);
router.post('/product-types', addProductType);
router.get('/product-types', getProductTypes);
router.delete('/product-types/:productType', deleteProductType);
router.put('/products/:tableName/:id', upload.array('images'), updateProduct);
router.delete('/products/:tableName/:id', deleteProduct);
router.patch('/products/:tableName/:id/toggle-status', toggleProductStatus);
router.patch('/products/:tableName/:id/toggle-fast-running', toggleFastRunning);
router.patch('/products/:tableName/:id/toggle-free', toggleFree);
router.post('/products/bulk-upload-images', upload.array('images', 200), bulkUploadProductImages);
router.patch('/products/bulk-discount', bulkUpdateDiscount);

// Image proxy for downloading product images without CORS issues
router.get('/proxy-image', async (req, res) => {
  try {
    const { url } = req.query;
    if (!url) return res.status(400).send('URL required');
    const response = await axios.get(url, { responseType: 'arraybuffer' });
    const contentType = response.headers['content-type'] || 'image/jpeg';
    res.set('Content-Type', contentType);
    res.send(response.data);
  } catch (err) {
    res.status(500).send(err.message);
  }
});

// Brands CRUD
router.get('/brands', async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM public.brands ORDER BY brand_name ASC');
    res.json(result.rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.post('/brands', async (req, res) => {
  const { brand_name } = req.body;
  if (!brand_name || !brand_name.trim()) return res.status(400).json({ error: 'brand_name is required' });
  try {
    const result = await pool.query(
      'INSERT INTO public.brands (brand_name) VALUES ($1)',
      [brand_name.trim()]
    );
    if (result.rows.length === 0) return res.status(409).json({ error: 'Brand already exists' });
    res.json(result.rows[0]);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.delete('/brands/:id', async (req, res) => {
  try {
    await pool.query('DELETE FROM public.brands WHERE id = $1', [req.params.id]);
    res.json({ message: 'Brand deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;