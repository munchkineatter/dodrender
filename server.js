const express = require('express');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS for Excel requests
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

// Store game state in memory
let gameState = {
    eliminatedPrizes: [],
    lastUpdated: new Date().toISOString()
};

// API: Get current game state
app.get('/api/gamestate', (req, res) => {
    res.json(gameState);
});

// API: Update game state (called from Excel)
app.post('/api/update', (req, res) => {
    try {
        const { eliminatedPrizes } = req.body;
        
        if (Array.isArray(eliminatedPrizes)) {
            gameState.eliminatedPrizes = eliminatedPrizes;
            gameState.lastUpdated = new Date().toISOString();
            
            console.log(`Game state updated: ${eliminatedPrizes.length} prizes eliminated`);
            res.json({ success: true, message: 'Game state updated' });
        } else {
            res.status(400).json({ success: false, message: 'Invalid data format' });
        }
    } catch (error) {
        console.error('Error updating game state:', error);
        res.status(500).json({ success: false, message: 'Server error' });
    }
});

// API: Reset game state
app.post('/api/reset', (req, res) => {
    gameState = {
        eliminatedPrizes: [],
        lastUpdated: new Date().toISOString()
    };
    console.log('Game state reset');
    res.json({ success: true, message: 'Game state reset' });
});

// API: Add single eliminated prize (alternative endpoint)
app.post('/api/eliminate', (req, res) => {
    try {
        const { prize } = req.body;
        
        if (prize && !gameState.eliminatedPrizes.includes(Number(prize))) {
            gameState.eliminatedPrizes.push(Number(prize));
            gameState.lastUpdated = new Date().toISOString();
            
            console.log(`Prize eliminated: $${prize}`);
            res.json({ success: true, message: `Prize $${prize} eliminated` });
        } else {
            res.json({ success: true, message: 'Prize already eliminated or invalid' });
        }
    } catch (error) {
        console.error('Error eliminating prize:', error);
        res.status(500).json({ success: false, message: 'Server error' });
    }
});

// Health check for Render
app.get('/health', (req, res) => {
    res.json({ status: 'ok', uptime: process.uptime() });
});

// Serve the main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
    console.log(`🎄 Prize Display Server running on port ${PORT}`);
    console.log(`   Open http://localhost:${PORT} to view`);
});

