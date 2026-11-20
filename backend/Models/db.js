const mongoose = require('mongoose');

const mongoUrl = process.env.MONGO_CONN;

if (!mongoUrl) {
    console.error('MongoDB Connection Error: MONGO_CONN is not configured. Add it to backend/.env.');
} else {
    mongoose.connect(mongoUrl)
        .then(() => {
            console.log('MongoDB Connected...');
        }).catch((err) => {
            console.error('MongoDB Connection Error:', err.message);
        });
}