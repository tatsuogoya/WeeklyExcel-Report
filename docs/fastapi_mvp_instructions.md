# Weekly Excel Report – FastAPI MVP Implementation Instructions

This document provides a code-free implementation guide for building
the Weekly Excel Report Generator as a FastAPI-based MVP.
It is based on the project README.md and serves as instructions
for implementation using VS Code + GitHub Copilot.


MVP Core Objectives
Remote Control: Start and stop activity simulation via HTTP requests.
State Management: Real-time status reporting (Running, Paused, Idle).
Dynamic Configuration: Allow changing presets and intervals through the API.
Technical Architecture
1. Backend Framework
FastAPI: Main web framework for the REST API.
Uvicorn: ASGI server for running the FastAPI application.
BackgroundTasks: Utilize FastAPI's BackgroundTasks or a dedicated threading model to run the activity simulation without blocking the API response.
2. State & Persistence
Runtime State: Use a singleton or a global state object to track the current simulation status and configuration.
File-based Config: Continue using activity_config.json for persistence, but allow the API to overwrite it on-the-fly.
Proposed API Endpoints
[Control Endpoints]
POST /v1/control/start: Initiates the activity simulation.
POST /v1/control/stop: Gracefully terminates the simulation.
PATCH /v1/control/preset/{preset_name}: Switches the current operation mode (stealth, aggressive, etc.).
[Status & Config Endpoints]
GET /v1/status: Returns current activity state, uptime, and last action performed.
GET /v1/config: Retrieves the active configuration settings.
PUT /v1/config: Updates specific parameters like activity_interval or mouse_move_distance.
Implementation Strategy
Phase 1: Core Engine Adaptation
Refactor existing activity_keeper.py logic into a reusable class/module that can be instantiated by the FastAPI app.
Ensure the simulation loop is non-blocking and can be interrupted programmatically.
Phase 2: API Layer Implementation
Define Pydantic models for configuration and status responses.
Implement the endpoints using dependency injection for the state manager.
Phase 3: Background Task Management
Implement a mechanism to prevent multiple simulation instances from running simultaneously.
Add lifecycle hooks (startup, shutdown) to ensure resources are cleaned up.
IMPORTANT

Since this MVP involves keyboard and mouse automation, the FastAPI server must run on the same machine that handles the Teams status. It cannot be containerized (Docker) without special configuration for GUI access.

TIP

For the initial MVP, focus on the stealth and standard presets as they provide the most value with the least risk of detection.
