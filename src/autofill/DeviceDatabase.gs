/**
 * Smart Home Inventory - Device Database
 * 
 * This file contains the device information map and related utility functions.
 * Separated from the main script for better organization and maintainability.
 * 
 * @author mariobob
 * @version 1.0
 */

// =============================================================================
// DEVICE INFORMATION MAP
// =============================================================================
// Key: Model number
// Value: Object with device properties (Device type, Manufacturer, etc.)
// Ordered by: Device type, then Actual device name
// =============================================================================

const DEVICE_MAP = {
  // Light > Hue color lamp
  "9290012573A": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue color lamp",
    "Zigbee ID": "LCT015",
    "Luminous flux": "806 lm",
    "Power": "9.5 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  "9290022168": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue color lamp",
    "Zigbee ID": "LCA001",
    "Luminous flux": "806 lm",
    "Power": "9 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  "9290024688": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue color lamp",
    "Zigbee ID": "LCA006",
    "Luminous flux": "1055 lm",
    "Power": "11 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  "9290024688A": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue color lamp",
    "Zigbee ID": "LCA006",
    "Luminous flux": "1055 lm",
    "Power": "11 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  "9290024716": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue color lamp",
    "Zigbee ID": "LCA008",
    "Luminous flux": "1600 lm",
    "Power": "15 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  
  // Light > Hue color candle
  "9290022942": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue color candle",
    "Zigbee ID": "LCE002",
    "Luminous flux": "470 lm",
    "Power": "5.3 W",
    "Plug type": "E14",
    "Smart": "Yes"
  },
  
  // Light > Hue color spot
  "9290019531": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue color spot",
    "Zigbee ID": "LCG002",
    "Luminous flux": "350 lm",
    "Power": "5.7 W",
    "Plug type": "GU10",
    "Smart": "Yes"
  },
  
  // Light > Hue white lamp
  "9290024692": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue white lamp",
    "Zigbee ID": "LWA017",
    "Luminous flux": "1055 lm",
    "Power": "9.5 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  "9290024406": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue white lamp",
    "Zigbee ID": "LWU001",
    "Luminous flux": "470 lm",
    "Power": "5.7 W",
    "Plug type": "E14",
    "Smart": "Yes"
  },
  "8718696449578": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue white lamp",
    "Zigbee ID": "",
    "Luminous flux": "806 lm",
    "Power": "9 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  "9290018216": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue white lamp",
    "Zigbee ID": "",
    "Luminous flux": "806 lm",
    "Power": "9 W",
    "Plug type": "E27",
    "Smart": "Yes"
  },
  
  // Light > Hue white candle
  "9290020399": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue white candle",
    "Zigbee ID": "LWE002",
    "Luminous flux": "470 lm",
    "Power": "5.5 W",
    "Plug type": "E14",
    "Smart": "Yes"
  },
  
  // Light > Hue Go
  "915005822001": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue Go",
    "Zigbee ID": "LCT026",
    "Luminous flux": "530 lm",
    "Power": "6 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  "915005821901": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue Go",
    "Zigbee ID": "7602031P7",
    "Luminous flux": "530 lm",
    "Power": "6.2 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Light > Hue Play
  "915005734101": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue Play",
    "Zigbee ID": "440400982842",
    "Luminous flux": "500 lm",
    "Power": "6.7 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  "915005733901": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue Play",
    "Zigbee ID": "440400982842",
    "Luminous flux": "500 lm",
    "Power": "6.7 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  "915005733701": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue Play",
    "Zigbee ID": "440400982841",
    "Luminous flux": "500 lm",
    "Power": "6.6 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Light > Hue Infuse ceiling
  "915005997301": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue Infuse ceiling",
    "Zigbee ID": "915005997301",
    "Luminous flux": "2350 lm",
    "Power": "42 W",
    "Plug type": "AC",
    "Smart": "Yes"
  },
  
  // Light > Signe gradient table
  "915005987001": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Signe gradient table",
    "Zigbee ID": "915005987001",
    "Luminous flux": "1040 lm",
    "Power": "11.8 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  "915005986901": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Signe gradient table",
    "Zigbee ID": "915005986901",
    "Luminous flux": "1040 lm",
    "Power": "11.8 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Light > Hue gradient lightstrip
  "9290029949": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue gradient lightstrip",
    "Zigbee ID": "LCX004",
    "Luminous flux": "1800 lm",
    "Power": "20 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Light > Hue lightstrip plus
  "929002269101": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue lightstrip plus",
    "Zigbee ID": "LCL001",
    "Luminous flux": "1600 lm",
    "Power": "20 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Light > Hue lightstrip outdoor
  "9290022891": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue lightstrip outdoor",
    "Zigbee ID": "LCL003",
    "Luminous flux": "1650 lm",
    "Power": "39.5 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Light > Hue Amarant wall washer
  "915005843401": {
    "Device type": "Light",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue Amarant wall washer",
    "Zigbee ID": "1746630P7",
    "Luminous flux": "1420 lm",
    "Power": "20 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Light > Yeelight lightstrip plus
  // (No model numbers in inventory)
  
  // Light > Mi LED Desk Lamp 1S
  "MJTD01SYL": {
    "Device type": "Light",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Mi LED Desk Lamp 1S",
    "Zigbee ID": "-",
    "Luminous flux": "520 lm",
    "Power": "7.5 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Plug > Hue smart plug
  "9290022404": {
    "Device type": "Plug",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue smart plug",
    "Zigbee ID": "LOM001",
    "Luminous flux": "N/A",
    "Power": "2200 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  "9290030506": {
    "Device type": "Plug",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue smart plug",
    "Zigbee ID": "LOM007",
    "Luminous flux": "N/A",
    "Power": "2300 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Plug > Mi Smart Plug
  "ZNCZ04LM": {
    "Device type": "Plug",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Mi Smart Plug",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "2300 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Sensor > Hue motion sensor
  "9290030675": {
    "Device type": "Sensor",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue motion sensor",
    "Zigbee ID": "SML003",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "AAA",
    "Smart": "Yes"
  },
  "9290012607": {
    "Device type": "Sensor",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue motion sensor",
    "Zigbee ID": "SML001",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "AAA",
    "Smart": "Yes"
  },
  
  // Sensor > Hue motion sensor outdoor
  "9290030674": {
    "Device type": "Sensor",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue motion sensor outdoor",
    "Zigbee ID": "SML004",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "AA",
    "Smart": "Yes"
  },
  
  // Sensor > Mi motion sensor
  // (No model numbers in inventory)
  
  // Switch > Hue dimmer switch V1
  "324131137411": {
    "Device type": "Switch",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue dimmer switch V1",
    "Zigbee ID": "RWL021",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "CR2450",
    "Smart": "Yes"
  },
  
  // Switch > Hue dimmer switch V2
  "9290023986": {
    "Device type": "Switch",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue dimmer switch V2",
    "Zigbee ID": "RWL022",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "CR2450",
    "Smart": "Yes"
  },
  
  // Switch > Hue switch module
  "9290030171": {
    "Device type": "Switch",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue switch module",
    "Zigbee ID": "",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "CR2450",
    "Smart": "Yes"
  },
  
  // Switch > Hue smart button
  "9290022230": {
    "Device type": "Switch",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue smart button",
    "Zigbee ID": "ROM001",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "CR2032",
    "Smart": "Yes"
  },
  
  // Switch > Hue tap dial switch
  "9290035001": {
    "Device type": "Switch",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue tap dial switch",
    "Zigbee ID": "RDM002",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "CR2032",
    "Smart": "Yes"
  },
  
  // Speaker > Nest Mini
  "H2C": {
    "Device type": "Speaker",
    "Manufacturer": "Google",
    "Actual device name": "Nest Mini",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "15.4 W",
    "Plug type": "UK",
    "Smart": "Yes"
  },
  
  // Speaker > Nest Audio
  "GXCA6": {
    "Device type": "Speaker",
    "Manufacturer": "Google",
    "Actual device name": "Nest Audio",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "30 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Speaker > Nest Hub
  "GUIK2": {
    "Device type": "Speaker",
    "Manufacturer": "Google",
    "Actual device name": "Nest Hub",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "15 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Speaker > Echo Dot
  "O78MP8": {
    "Device type": "Speaker",
    "Manufacturer": "Amazon",
    "Actual device name": "Echo Dot",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "15 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Speaker > Echo Dot with Clock
  "B7W644": {
    "Device type": "Speaker",
    "Manufacturer": "Amazon",
    "Actual device name": "Echo Dot with Clock",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "15 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  "C4E8S3": {
    "Device type": "Speaker",
    "Manufacturer": "Amazon",
    "Actual device name": "Echo Dot with Clock",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "15 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Wifi > Google Wifi
  "GJ2CQ": {
    "Device type": "Wifi",
    "Manufacturer": "Google",
    "Actual device name": "Google Wifi",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "15.4 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Wifi > Nest Wifi Pro
  "G6ZUC": {
    "Device type": "Wifi",
    "Manufacturer": "Google",
    "Actual device name": "Nest Wifi Pro",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "22.5 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Wifi > ROG Rapture GT-AXE16000
  "90IG06W0-MU2A10": {
    "Device type": "Wifi",
    "Manufacturer": "ASUS",
    "Actual device name": "ROG Rapture GT-AXE16000",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "65 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Wifi > GT-AXE11000
  "90IG06E0-MO1R00": {
    "Device type": "Wifi",
    "Manufacturer": "ASUS",
    "Actual device name": "ROG Rapture GT-AXE11000",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "65 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Climate > Mi Air Purifier Pro
  "AC-M3-CA": {
    "Device type": "Climate",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Mi Air Purifier Pro",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "66 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Climate > Mi Smart Antibacterial Humidifier
  "ZNJSQ01DEM": {
    "Device type": "Climate",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Mi Smart Antibacterial Humidifier",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "25 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Climate > Mi Smart Standing Fan Pro
  "ZLBPSP01XY": {
    "Device type": "Climate",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Mi Smart Standing Fan Pro",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "24 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Climate > Mi Smart Space Heater S
  "KRDNQ03ZM": {
    "Device type": "Climate",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Mi Smart Space Heater S",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "2200 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },

  // Climate > Xiaomi Smart Tower Heater Lite
  "LSNFJ02LX": {
    "Device type": "Climate",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Xiaomi Smart Tower Heater Lite",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "2000 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },

  // Climate > Bergen AION AI 12K
  "BER12AUCXB-K6DNA1A/I": {
    "Device type": "Climate",
    "Manufacturer": "Bergen",
    "Actual device name": "Bergen AION AI 12K",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "3500 W",
    "Plug type": "AC",
    "Smart": "Yes"
  },
  
  // Climate > Terma Home Comfort thermostat
  "WBR3": {
    "Device type": "Climate",
    "Manufacturer": "Terma",
    "Actual device name": "Terma Home Comfort thermostat",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "",
    "Plug type": "AC",
    "Smart": "Yes"
  },
  
  // Camera > Nest Cam wired
  "GJQ9T": {
    "Device type": "Camera",
    "Manufacturer": "Google",
    "Actual device name": "Nest Cam wired",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "7.5 W",
    "Plug type": "USB",
    "Smart": "Yes"
  },
  
  // Camera > Mi 360 Home Security Camera 2K Pro
  "MJSXJ06CM": {
    "Device type": "Camera",
    "Manufacturer": "Xiaomi",
    "Actual device name": "Mi 360° Home Security Camera 2K Pro",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "10 W",
    "Plug type": "microUSB",
    "Smart": "Yes"
  },
  
  // Camera > Blink Mini
  "BCM00300U": {
    "Device type": "Camera",
    "Manufacturer": "Amazon",
    "Actual device name": "Blink Mini",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "5 W",
    "Plug type": "microUSB",
    "Smart": "Yes"
  },
  
  // Camera > Blink Outdoor
  "BCM01400U": {
    "Device type": "Camera",
    "Manufacturer": "Amazon",
    "Actual device name": "Blink Outdoor",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "Battery",
    "Plug type": "AA",
    "Smart": "Yes"
  },
  
  // TV > Hue sync box
  "9290022758": {
    "Device type": "TV",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Hue sync box",
    "Zigbee ID": "-",
    "Luminous flux": "N/A",
    "Power": "7.3 W",
    "Plug type": "EU",
    "Smart": "Yes"
  },
  
  // Hub > Philips hue bridge
  "3241312018A": {
    "Device type": "Hub",
    "Manufacturer": "Philips Hue",
    "Actual device name": "Philips hue bridge",
    "Zigbee ID": "BSB002",
    "Luminous flux": "N/A",
    "Power": "5 W",
    "Plug type": "EU",
    "Smart": "Yes"
  }
};

// =============================================================================
// DATABASE UTILITY FUNCTIONS
// =============================================================================

/**
 * Gets device information for a specific model.
 * 
 * @param {string} model - The model number to look up
 * @return {Object|null} Device information or null if not found
 */
function getDeviceInfo(model) {
  if (!model) return null;
  // Check hardcoded database first
  if (DEVICE_MAP[model]) return DEVICE_MAP[model];
  // Fall back to user-added custom devices (stored in Script Properties)
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('CUSTOM_DEVICES');
    if (raw) {
      const custom = JSON.parse(raw);
      if (custom[model]) return custom[model];
    }
  } catch (e) {
    Logger.log('CustomDeviceDB read error: ' + e.message);
  }
  return null;
}

/**
 * Validates the device map for duplicates and inconsistencies.
 * Run this function manually to check data integrity.
 * 
 * @return {Object} Validation report with totalModels, duplicates, and isValid
 */
function validateDeviceMap() {
  const models = Object.keys(DEVICE_MAP);
  const duplicates = [];
  const modelSet = new Set();
  
  models.forEach(model => {
    if (modelSet.has(model)) {
      duplicates.push(model);
    }
    modelSet.add(model);
  });
  
  if (duplicates.length > 0) {
    Logger.log(`Found ${duplicates.length} duplicate model(s): ${duplicates.join(", ")}`);
  } else {
    Logger.log(`Device map validated successfully. Total models: ${models.length}`);
  }
  
  return {
    totalModels: models.length,
    duplicates: duplicates,
    isValid: duplicates.length === 0
  };
}

/**
 * Lists all models in the device map.
 * Useful for debugging and verification.
 */
function listAllModels() {
  const models = Object.keys(DEVICE_MAP);
  models.sort();
  
  Logger.log(`Total models in database: ${models.length}`);
  models.forEach((model, index) => {
    const device = DEVICE_MAP[model];
    Logger.log(`${index + 1}. ${model} - ${device["Manufacturer"]} ${device["Actual device name"]}`);
  });
}

/**
 * Gets statistics about the device database.
 * 
 * @return {Object} Statistics including device type and manufacturer counts
 */
function getDeviceDatabaseStats() {
  const models = Object.keys(DEVICE_MAP);
  const deviceTypes = {};
  const manufacturers = {};
  
  models.forEach(model => {
    const device = DEVICE_MAP[model];
    const type = device["Device type"];
    const manufacturer = device["Manufacturer"];
    
    // Count by device type
    deviceTypes[type] = (deviceTypes[type] || 0) + 1;
    
    // Count by manufacturer
    manufacturers[manufacturer] = (manufacturers[manufacturer] || 0) + 1;
  });
  
  return {
    totalModels: models.length,
    deviceTypes: deviceTypes,
    manufacturers: manufacturers
  };
}
