// approvalCard.ts

export function getApprovalAdaptiveCard(data: any, locale: string = 'en'): any {
  const translations: any = {
    en: {
      title: "Action Required: Approval Needed",
      prompt: "Please review the following request and choose to Proceed or Cancel.",
      approve: "✅ Proceed",
      reject: "❌ Cancel",
      help: "If you have any questions, please contact your administrator.",
      inputLabel: "Additional Comments:",
      inputPlaceholder: "Enter your comments here..."
    },
    fr: {
      title: "Action requise : Approbation nécessaire",
      prompt: "Veuillez examiner la demande suivante et choisir de procéder ou d'annuler.",
      approve: "✅ Procéder",
      reject: "❌ Annuler",
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
// Define a union of all AdaptiveCard element types you want to support
type AdaptiveCardElement =
  | {
      type: "Container";
      items: {
        type: "TextBlock";
        text: string;
        wrap: boolean;
        spacing: string;
      }[];
      style: string;
    }
  | {
      type: "TextBlock";
      text: string;
      wrap: boolean;
      spacing: string;
    }
  | {
      type: "Input.Text";
      id: string;
      value: string;
      placeholder: string;
      isMultiline: boolean;
      spacing: string;
    };

// Force parameterInputs to be AdaptiveCardElement[]
const parameterInputs: AdaptiveCardElement[] = Object.entries(parameters)
  .filter(([_, value]) => value !== undefined)
  .flatMap(([key, value]): AdaptiveCardElement[] => {
    if (typeof value === "object" && Array.isArray(value)) {
      // array branch → Containers
      return value.map((item, index) => ({
        type: "Container" as const,
        items: [
          {
            type: "TextBlock" as const,
            text: `${key.replace(/_/g, " ")} - ${index}:`,
            wrap: true,
            spacing: "Medium",
          },
          {
            type: "TextBlock" as const,
            text: `${item.Name} : ${item.Type}`,
            wrap: true,
            spacing: "Medium",
          }
        ],
        style: "emphasis",
      }));
    }

    // primitive branch → TextBlock + Input.Text
    return [
      {
        type: "TextBlock" as const,
        text: `${key.replace(/_/g, " ")}:`,
        wrap: true,
        spacing: "Medium"
      },
      {
        type: "Input.Text" as const,
        id: key,
        value: String(value),
        placeholder: `Enter ${key.replace(/_/g, " ")}`,
        isMultiline: String(value).length > 50,
        spacing: "Small"
      },
    ];
  });


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
          "msteams": { "type": "messageBack", "text": "Proceed" }
        }
      },
      {
        "type": "Action.Submit",
        "title": t.reject,
        "style": "destructive",
        "data": {
          "msteams": { "type": "messageBack", "text": "Cancel" }
        }
      }
    ]
  };
}