const fs = require('fs');

const filePath = './config/package-solution.json'; // Update with actual JSON file path

function decrementVersion(version) {

    // Special case: do nothing if version is 1.0.0.0
    if ( version === "1.0.0.0") return version

    // Special case: do nothing if version is 0.0.0.0
    if ( version === "0.0.0.0") return version

    const parts = version.split('.').map(part => parseInt(part, 10));

    for (let i = parts.length - 1; i >= 0; i--) {
        if (parts[i] > 0) {
            parts[i]--;
            break;
        } else {
            parts[i] = 9; // Set the current part to 9 if it's 0
            if (i === 0) {
                // Prevent setting to 9 at the highest order part
                parts[i] = 0;
                break;
            }
        }
    }

    return parts.join('.');
}

function updateVersion() {
    try {
        const json = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        console.log("Current version:", json.solution.version);
        json.solution.version = decrementVersion(json.solution.version);
        fs.writeFileSync(filePath, JSON.stringify(json, json.solution.version, 2), 'utf8');
        console.log(`Version updated: ${json.solution.version}`);
    } catch (error) {
        console.error("Error updating version:", error);
    }
}

updateVersion();
