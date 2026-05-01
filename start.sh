#!/bin/bash

# Export Fastmail API key (already embedded in code, but can be overridden)
export FASTMAIL_API_KEY="${FASTMAIL_API_KEY:-fmu1-d01e43a8-f3ca5b7579eb5aac1f2df46b23440060-0-e12f16ead889af9da7e5dbf2720a49ff}"

# Port configuration (default: 3000)
export PORT="${PORT:-3000}"

# Start the server
node index.js
