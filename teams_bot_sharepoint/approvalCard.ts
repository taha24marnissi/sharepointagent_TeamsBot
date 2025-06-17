// approvalCard.ts

export function getApprovalAdaptiveCard(data: any, locale: string = 'en'): any {
  const translations: any = {
    en: {
      title: "Action Required: Approval Needed",
      prompt: "Please review the following request and choose to Approve or Reject.",
      approve: "✅ Approve",
      reject: "❌ Reject",
      help: "If you have any questions, please contact your administrator."
    },
    fr: {
      title: "Action requise : Approbation nécessaire",
      prompt: "Veuillez examiner la demande suivante et choisir d'approuver ou de rejeter.",
      approve: "✅ Approuver",
      reject: "❌ Rejeter",
      help: "Si vous avez des questions, veuillez contacter votre administrateur."
    }
  };
  const t = translations[locale] || translations['en'];
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
            "text": data.response ?? (locale === 'fr' ? "Aucune réponse de l'agent." : "No response from the agent."),
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
          ...data,
          "msteams": { "type": "messageBack", "text": "ApproveButton" }
        }
      },
      {
        "type": "Action.Submit",
        "title": t.reject,
        "style": "destructive",
        "data": {
          ...data,
          "msteams": { "type": "messageBack", "text": "RejectButton" }
        }
      }
    ]
  };
}
