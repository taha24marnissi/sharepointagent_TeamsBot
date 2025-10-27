/**
 * Creates a professional loading/progress Adaptive Card with ProgressBar
 * Shows visual progress indicator and status message
 */
export function createLoadingCard(message?: string, elapsed?: number): any {
  const displayMessage = message || "Processing your request...";
  return {
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Your request is being processed",
        wrap: true,
        size: "Large",
        weight: "Bolder"
      },
      {
        type: "ProgressBar"
      },
      {
        type: "TextBlock",
        text: displayMessage,
        spacing: "ExtraSmall",
        size: "Small",
        wrap: true
      },
      ...(elapsed ? [
        {
          type: "TextBlock",
          text: `⏱️ Time elapsed: ${elapsed} seconds`,
          spacing: "Small",
          size: "Small",
          color: "Attention",
          wrap: true
        }
      ] : [])
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
  };
}

/**
 * Creates a completion card with green progress bar at 100%
 * Shows success status with checkmark icon
 */
export function createCompletionCard(duration: number): any {
  return {
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Your request completed",
        wrap: true,
        size: "Large",
        weight: "Bolder"
      },
      {
        type: "ProgressBar",
        value: 100,
        color: "Good"
      },
      {
        type: "ColumnSet",
        columns: [
          {
            type: "Column",
            width: "auto",
            items: [
              {
                type: "Icon",
                name: "CheckmarkCircle",
                size: "Small",
                color: "Good"
              }
            ],
            verticalContentAlignment: "Center"
          },
          {
            type: "Column",
            width: "stretch",
            items: [
              {
                type: "TextBlock",
                text: `The request was successfully processed in ${duration} seconds`,
                spacing: "ExtraSmall",
                size: "Small",
                color: "Good",
                wrap: true
              }
            ],
            verticalContentAlignment: "Center",
            spacing: "Small"
          }
        ],
        spacing: "ExtraSmall"
      }
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
  };
}
