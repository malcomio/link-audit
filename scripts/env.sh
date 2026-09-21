#!/bin/bash

# Copy the .env.example file if .env does not already exist.
if ! [[ -f .env ]]; then
  cp .env.example .env
fi
