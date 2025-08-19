// approvalCard.ts

export function getApprovalAdaptiveCard(data: any, locale: string = 'en'): any {
  const translations: any = {
    en: {
      title: "Action Required: Approval Needed",
      prompt: "Please review the following request and choose to Approve or Reject.",
      approve: "✅ Approve",
      reject: "❌ Reject",
      help: "If you have any questions, please contact your administrator.",
      inputLabel: "Additional Comments:",
      inputPlaceholder: "Enter your comments here..."
    },
    fr: {
      title: "Action requise : Approbation nécessaire",
      prompt: "Veuillez examiner la demande suivante et choisir d'approuver ou de rejeter.",
      approve: "✅ Approuver",
      reject: "❌ Rejeter",
      help: "Si vous avez des questions, veuillez contacter votre administrateur.",
      inputLabel: "Commentaires supplémentaires :",
      inputPlaceholder: "Entrez vos commentaires ici..."
    }
  };
  const t = translations[locale] || translations['en'];
  
  // Parse response and create inputs
  const responseObj = JSON.parse(data.response.replace(/'/g, '"').replace(/True/g, 'true')
      .replace(/False/g, 'false'));
  const parameters = responseObj.parameters || {};
  
  // Add validation to ensure parameters is an object
  if (typeof parameters !== 'object' || parameters === null) {
    throw new Error('Parameters must be an object');
  }

  const parameterInputs = Object.entries(parameters)
    .filter(([_, value]) => value !== undefined) // Filter out undefined values
    .map(([key, value]) => [
      {
        "type": "TextBlock",
        "text": `${key.replace(/_/g, ' ').toUpperCase()}:`,
        "wrap": true,
        "spacing": "Medium"
      },
      {
        "type": "Input.Text",
        "id": key,
        "value": typeof value === 'object' ? JSON.stringify(value) : String(value), // Handle objects
        "placeholder": `Enter ${key.replace(/_/g, ' ')}`,
        "isMultiline": typeof value === 'object', // Make input multiline for objects
        "spacing": "Small"
      }
    ]).flat();

  return {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      {
        "type": "ColumnSet",
        "columns": [
          {
            "type": "Column",
            "width": "auto",
            "items": [
              {
                "type": "Image",
                "url": "https://www.clipartmax.com/png/middle/118-1180913_approve-document-icons-tick-and-cross-icon.png",
                "size": "Small",
                "style": "Person"
              }
            ]
          },
          {
            "type": "Column",
            "width": "stretch",
            "items": [
              {
                "type": "TextBlock",
                "text": t.title,
                "weight": "Bolder",
                "size": "Large"
              },
              {
                "type": "TextBlock",
                "text": t.prompt,
                "isSubtle": true,
                "wrap": true
              }
            ]
          }
        ]
      },
      {
        "type": "Container",
        "items": [
          {
            "type": "TextBlock",
            "text": responseObj.text,
            "wrap": true,
            "spacing": "Medium",
            "size": "Medium"
          }
        ],
        "style": "emphasis",
        "bleed": true
      },
      ...parameterInputs,
      {
        "type": "TextBlock",
        "text": t.help,
        "isSubtle": true,
        "wrap": true,
        "spacing": "Medium"
      }
    ],
    "actions": [
      {
        "type": "Action.Submit",
        "title": t.approve,
        "style": "positive",
        "data": {
          "msteams": { "type": "messageBack", "text": "ApproveButton" }
        }
      },
      {
        "type": "Action.Submit",
        "title": t.reject,
        "style": "destructive",
        "data": {
          "msteams": { "type": "messageBack", "text": "RejectButton" }
        }
      }
    ]
  };
}
