#!/bin/bash
echo "Aurora deployment starting..."
cd "$DEPLOYMENT_TARGET"
npm install --production
echo "Aurora deployment complete"
