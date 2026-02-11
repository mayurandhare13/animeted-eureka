/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Add event listeners for simplified API buttons
    document.getElementById("check-permission").onclick = checkPermissionStatus;
    document.getElementById("request-permission").onclick = requestMicrophonePermission;
    document.getElementById("get-devices").onclick = getMicrophoneDevices;
  }
});

let hasShownPermissionHint = false;

// ============================================
// 1. Check Permission Status
// ============================================
export async function checkPermissionStatus() {
  try {
    console.log("Checking permission status...");
    updateStatus("Checking microphone permission status...", "info");

    // Check if API is available
    if (!Office.devicePermission) {
      updateStatus("❌ Device Permission API is not available", "error");
      updatePermissionType("Requires Office on the web (Excel, PowerPoint, Word, or Outlook)");
      return;
    }

    // Check platform
    if (Office.context.platform !== Office.PlatformType.OfficeOnline) {
      updateStatus("❌ Not running on Office on the web", "error");
      updatePermissionType("Current platform: " + Office.context.platform);
      return;
    }

    // Check browser permission state using Web API
    if (navigator.permissions) {
      try {
        const permissionStatus = await navigator.permissions.query({ name: 'microphone' as PermissionName });
        console.log("Permission state:", permissionStatus.state);

        if (permissionStatus.state === 'granted') {
          updateStatus("✅ Microphone permission is GRANTED", "success");
          updatePermissionType("Permission is active. Type: Cannot determine if 'Always' or 'Allow Once' from browser API alone.");
          hidePermissionHint();
        } else if (permissionStatus.state === 'denied') {
          updateStatus("⛔ Microphone permission is DENIED", "error");
          updatePermissionType("User has blocked microphone access. Clear site permissions in browser settings.");
          hidePermissionHint();
        } else {
          updateStatus("⚠️ Microphone permission is PROMPT (not yet decided)", "info");
          updatePermissionType("User hasn't granted or denied permission yet. Click 'Request Permission' to prompt.");
          showPermissionHintOnce("Tip: Click 'Request Microphone Permission' to trigger the browser prompt before trying to list devices.");
        }

        // Listen for permission changes
        permissionStatus.onchange = () => {
          console.log("Permission changed to:", permissionStatus.state);
          checkPermissionStatus(); // Recheck when permission changes
        };
      } catch (error) {
        console.error("Could not query permissions:", error);
        updateStatus("✅ API is available, but cannot check permission state", "info");
        updatePermissionType("Click 'Request Permission' to proceed.");
      }
    } else {
      updateStatus("✅ API is available", "success");
      updatePermissionType("navigator.permissions API not available. Click 'Request Permission' to proceed.");
    }
  } catch (error) {
    console.error("Check permission error:", error);
    updateStatus(`Error checking permission: ${error instanceof Error ? error.message : error}`, "error");
  }
}

// ============================================
// 2. Request Microphone Permission
// ============================================
export async function requestMicrophonePermission() {
  try {
    console.log("Requesting microphone permission...");
    updateStatus("Requesting microphone permission...", "info");
    updatePermissionType("");

    // Check if API is available
    if (!Office.devicePermission) {
      updateStatus("❌ Device Permission API is not available", "error");
      updatePermissionType("Requires Office on the web");
      return;
    }

    // Check platform
    if (Office.context.platform !== Office.PlatformType.OfficeOnline) {
      updateStatus("❌ Only supported in Office on the web", "error");
      updatePermissionType("Current platform: " + Office.context.platform);
      return;
    }

    // Request permission
    const deviceCapabilities = [Office.DevicePermissionType.microphone];
    const isFirstTimeGrant = await Office.devicePermission.requestPermissions(deviceCapabilities);

    if (isFirstTimeGrant) {
      // User clicked "Allow" or "Allow Once" for the first time
      updateStatus("✅ Permission granted! Reloading in 2 seconds...", "success");
      updatePermissionType("First-time grant detected. The add-in will reload to apply permissions.");

      hidePermissionHint();

      setTimeout(() => {
        console.log("Reloading after first-time permission grant...");
        location.reload();
      }, 2000);
    } else {
      // Permission was already granted previously
      updateStatus("✅ Permission already granted (using existing permission)", "success");
      updatePermissionType("This permission was granted earlier. Type: 'Always' if persistent after tab close, 'Allow Once' if requires re-grant after tab close.");

      hidePermissionHint();

      // Test that we can access microphone
      await waitForPermissionPropagation();
      testMicrophoneAccess();
    }
  } catch (error) {
    console.error("Request permission error:", error);
    const errorMessage = error instanceof Error ? error.message : String(error);

    if (errorMessage.includes("denied")) {
      updateStatus("⛔ Permission DENIED by user", "error");
      updatePermissionType("User clicked 'Deny'. You can request permission again, and the user will be prompted again.");
    } else {
      updateStatus(`❌ Error: ${errorMessage}`, "error");
      updatePermissionType("");
    }
  }
}

// ============================================
// 3. Get Microphone Devices
// ============================================
export async function getMicrophoneDevices() {
  try {
    console.log("Getting microphone devices...");
    updateStatus("Enumerating microphone devices...", "info");

    if (!navigator.mediaDevices) {
      updateStatus("❌ MediaDevices API not available", "error");
      return;
    }

    // Get active stream first to expose device labels (W3C spec requirement)
    // This works cross-browser and is the recommended pattern per MDN
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    console.log("✅ Microphone access granted");

    // Enumerate devices while stream is active - labels will be populated
    const devices = await navigator.mediaDevices.enumerateDevices();
    const audioInputs = devices.filter(device => device.kind === 'audioinput');

    // Stop the stream immediately after enumeration
    stream.getTracks().forEach(track => track.stop());

    console.log(`Found ${audioInputs.length} audio input device(s)`);

    if (audioInputs.length === 0) {
      updateStatus("⚠️ No microphone devices found", "error");
      updatePermissionType("Make sure a microphone is connected.");
      clearDeviceList();
      return;
    }

    updateStatus(`✅ Found ${audioInputs.length} microphone device(s)`, "success");
    updatePermissionType("Devices are accessible with permission granted.");
    displayDevices(audioInputs, true);
  } catch (error) {
    console.error("Get devices error:", error);
    const errorMsg = error instanceof Error ? error.message : String(error);

    if (errorMsg.includes("Permission denied") || errorMsg.includes("NotAllowedError")) {
      updateStatus("⚠️ Microphone access denied. Click 'Request Permission' first.", "error");
      updatePermissionType("Permission not granted or user declined.");
    } else {
      updateStatus(`❌ Error: ${errorMsg}`, "error");
    }
    clearDeviceList();
  }
}

// ============================================
// Helper Functions
// ============================================

function updateStatus(message: string, type: "success" | "error" | "info") {
  const statusElement = document.getElementById("permission-status");
  const statusText = document.getElementById("status-text");

  if (statusElement && statusText) {
    statusText.textContent = message;
    statusElement.style.display = "block";

    // Update styling based on type
    if (type === "success") {
      statusElement.style.borderLeft = "4px solid #107C10";
      statusElement.style.background = "#DFF6DD";
    } else if (type === "error") {
      statusElement.style.borderLeft = "4px solid #D13438";
      statusElement.style.background = "#FDE7E9";
    } else {
      statusElement.style.borderLeft = "4px solid #0078D4";
      statusElement.style.background = "#F3F2F1";
    }
  }

  console.log(`[${type.toUpperCase()}] ${message}`);
}

function updatePermissionType(message: string) {
  const permissionTypeElement = document.getElementById("permission-type");
  if (permissionTypeElement) {
    permissionTypeElement.textContent = message;
    permissionTypeElement.style.display = message ? "block" : "none";
  }
}

function displayDevices(devices: MediaDeviceInfo[], hasLabels: boolean) {
  const deviceListContainer = document.getElementById("device-list");
  const devicesList = document.getElementById("devices");

  if (deviceListContainer && devicesList) {
    devicesList.innerHTML = '';

    devices.forEach((device, index) => {
      const li = document.createElement("li");
      li.style.padding = "8px";
      li.style.marginBottom = "4px";
      li.style.background = "#F3F2F1";
      li.style.borderRadius = "4px";

      const label = device.label || `(empty label)`;
      const deviceId = device.deviceId.substring(0, 20) + "...";

      li.innerHTML = `
        <div style="font-weight: 600;">${index + 1}. ${hasLabels ? '🎤' : '⚠️'} ${label}</div>
        <div style="font-size: 11px; color: #605E5C; margin-top: 4px;">ID: ${deviceId}</div>
      `;

      devicesList.appendChild(li);
    });

    deviceListContainer.style.display = "block";
  }
}

function clearDeviceList() {
  const deviceListContainer = document.getElementById("device-list");
  if (deviceListContainer) {
    deviceListContainer.style.display = "none";
  }
}

function showPermissionHintOnce(message: string) {
  if (hasShownPermissionHint) {
    return;
  }

  const hintElement = document.getElementById("permission-hint");
  if (hintElement) {
    hintElement.textContent = message;
    hintElement.style.display = "block";
    hasShownPermissionHint = true;
  }
}

function hidePermissionHint() {
  const hintElement = document.getElementById("permission-hint");
  if (hintElement) {
    hintElement.style.display = "none";
  }
}

async function testMicrophoneAccess() {
  try {
    console.log("Testing microphone access...");
    // await waitForPermissionPropagation();
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    console.log("✅ Microphone access successful:", stream);

    // Stop the stream immediately (we just wanted to test access)
    stream.getTracks().forEach(track => track.stop());
  } catch (error) {
    console.error("❌ Microphone access test failed:", error);
  }
}

function isFirefox() {
  return navigator.userAgent.toLowerCase().includes("firefox");
}

async function waitForPermissionPropagation() {
  if (!isFirefox()) {
    return;
  }

  // Firefox can apply iframe permissions policy asynchronously after load.
  await new Promise<void>((resolve) => {
    requestAnimationFrame(() => resolve());
  });
}
