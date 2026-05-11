const express = require('express');
const jwt = require('jsonwebtoken');
const Config = require('../models/Config');
const { authenticateToken } = require('../middleware/auth');

const router = express.Router();

const generateToken = (role) =>
  jwt.sign({ role }, process.env.JWT_SECRET, { expiresIn: '30d' });

router.post('/login', async (req, res) => {
  try {
    const { role, pin } = req.body;
    if (!role || !pin) {
      return res.status(400).json({ message: 'Role and PIN are required' });
    }
    if (!['owner', 'staff'].includes(role)) {
      return res.status(400).json({ message: 'Invalid role' });
    }

    let config = await Config.findOne();
    if (!config) config = await new Config().save();

    const validPin = role === 'owner' ? config.ownerPin : config.staffPin;
    if (pin !== validPin) {
      return res.status(401).json({ message: 'Invalid PIN' });
    }

    const token = generateToken(role);
    res.json({ token, role, cafeName: config.cafeName });
  } catch (error) {
    console.error('Login error:', error);
    res.status(500).json({ message: 'Login failed' });
  }
});

router.get('/verify', authenticateToken, async (req, res) => {
  try {
    const config = await Config.findOne();
    res.json({ valid: true, role: req.user.role, cafeName: config?.cafeName || 'The Sanctum Cafe' });
  } catch (error) {
    res.status(500).json({ message: 'Verification failed' });
  }
});

module.exports = router;
