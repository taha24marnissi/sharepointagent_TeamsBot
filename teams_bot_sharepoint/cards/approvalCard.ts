/**
 * Approval Card for SharePoint Operations
 * 
 * Creates an Adaptive Card for risky operations requiring user approval
 */

/**
 * Creates an approval Adaptive Card with warning styling
 * @param message - The approval message from the backend
 * @param disabled - Whether the buttons should be disabled (after user action)
 * @param selectedAction - The action that was selected ("Proceed" or "Cancel")
 * @returns Adaptive Card JSON object
 */
export function createApprovalCard(
  message: string, 
  disabled: boolean = false, 
  selectedAction?: string
): any {
  const card: any = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      {
        "type": "Container",
        "items": [
          {
            "type": "TextBlock",
            "text": disabled 
              ? (selectedAction === "Proceed" ? "✅ APPROVED" : "❌ CANCELLED")
              : "⚠️ APPROVAL REQUIRED",
            "weight": "Bolder",
            "size": "Large",
            "color": disabled 
              ? (selectedAction === "Proceed" ? "Good" : "Attention")
              : "Warning"
          }
        ]
      },
      {
        "type": "Container",
        "items": [
          {
            "type": "TextBlock",
            "text": message,
            "wrap": true,
            "spacing": "Medium"
          }
        ],
        "style": disabled 
          ? (selectedAction === "Proceed" ? "good" : "attention")
          : "warning",
        "bleed": true
      },
      {
        "type": "TextBlock",
        "text": disabled 
          ? `Action completed: ${selectedAction}`
          : "Please review the operation and choose an action:",
        "wrap": true,
        "spacing": "Medium",
        "weight": "Bolder"
      }
    ]
  };

  // Only add actions if not disabled
  if (!disabled) {
    card.actions = [
      {
        "type": "Action.Submit",
        "title": "✅ Proceed",
        "style": "positive",
        "data": {
          "action": "approval_proceed"
        }
      },
      {
        "type": "Action.Submit",
        "title": "❌ Cancel",
        "style": "destructive",
        "data": {
          "action": "approval_cancel"
        }
      }
    ];
  }

  return card;
}
