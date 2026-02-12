/* global document, Office, console, navigator */

type DialogAction = "request" | "devices";

let requestedAction: DialogAction = "request";

Office.onReady(() => {
  const actionButton = document.getElementById("action-button") as HTMLButtonElement | null;

  if (actionButton) {
    actionButton.onclick = () => runAction();
  }

  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    try {
      const message = JSON.parse(arg.message) as { type: string; action?: DialogAction };
      if (message.type === "start") {
        requestedAction = message.action === "devices" ? "devices" : "request";
        updateDialogLabel();
      }
    } catch (error) {
      console.error("Dialog message parse failed:", error);
    }
  });

  updateDialogLabel();
  Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
});

function updateDialogLabel() {
  const label = document.getElementById("action-label");
  const button = document.getElementById("action-button") as HTMLButtonElement | null;

  if (label) {
    label.textContent = requestedAction === "devices" ? "Enumerate microphone devices" : "Request microphone access";
  }

  if (button) {
    button.textContent = requestedAction === "devices" ? "Get devices" : "Request microphone";
  }
}

async function runAction() {
  const button = document.getElementById("action-button") as HTMLButtonElement | null;
  if (button) {
    button.disabled = true;
  }

  setStatus("Requesting microphone access...");

  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });

    if (requestedAction === "devices") {
      const devices = await navigator.mediaDevices.enumerateDevices();
      const audioInputs = devices.filter((device) => device.kind === "audioinput");

      stream.getTracks().forEach((track) => track.stop());

      const payload = audioInputs.map((device) => ({
        label: device.label || "(empty label)",
        deviceId: device.deviceId,
      }));

      Office.context.ui.messageParent(
        JSON.stringify({ type: "devices", devices: payload })
      );
    } else {
      stream.getTracks().forEach((track) => track.stop());
      Office.context.ui.messageParent(JSON.stringify({ type: "permission-granted" }));
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const denied = message.includes("NotAllowedError") || message.toLowerCase().includes("denied");

    Office.context.ui.messageParent(
      JSON.stringify({ type: denied ? "permission-denied" : "error", error: message })
    );
  } finally {
    if (button) {
      button.disabled = false;
    }
  }
}

function setStatus(message: string) {
  const status = document.getElementById("status");
  if (status) {
    status.textContent = message;
  }
}
