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
            "text": responseObj.message,
            "wrap": true,
            "spacing": "Medium",
            "size": "Medium"
          }
        ],
        "style": "emphasis",
        "bleed": true
      },      
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