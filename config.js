/**
 * Centralized configuration for test URLs
 * Change URLs here and they will apply everywhere in the project
 */

const config = {
  // Main application base URL
  baseURL: 'https://cloud.iqwater.de',
  
  // Login page URL
  loginURL: 'https://cloud.iqwater.de/anmeldung',
  
  // Profile page URL (for language switching)
  profileURL: 'https://cloud.iqwater.de/mein-profil',

  // Notification settings
  webhookUrl: process.env.WEBHOOK_URL || 'https://play.svix.com/in/CGZmAEQ0vsS0DakqdgS2442Pms1/',
};

module.exports = config;
