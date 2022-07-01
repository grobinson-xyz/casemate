const { CardFactory } = require('botbuilder');

exports.EmailCard = CardFactory.adaptiveCard({
    type: 'AdaptiveCard',
    body: [
        {
            type: 'TextBlock',
            size: 'Medium',
            weight: 'Bolder',
            text: 'The remote name could not be resolved - TrackingID#2206060040005687'
        },
        {
            type: 'ColumnSet',
            columns: [
                {
                    type: 'Column',
                    items: [
                        {
                            type: 'TextBlock',
                            weight: 'Bolder',
                            text: 'From: Gary Robinson; To: Reilly, Iain',
                            wrap: true
                        },
                        {
                            type: 'TextBlock',
                            spacing: 'None',
                            text: 'Created 06/29/2022',
                            isSubtle: true,
                            wrap: true
                        }
                    ],
                    width: 'stretch'
                }
            ]
        },
        {
            type: 'TextBlock',
            text: 'Hi Raja,\n \nMy name is Gary and I am from the App Services team. I can further assist you with troubleshooting your migrating your app service.  I have listed all of my backups and my managers contact information in my signature.\n\n',
            wrap: true
        },
        {
            type: 'TextBlock',
            text: 'Issue Description:',
            wrap: true,
            weight: 'Bolder',
            color: 'Accent'
        },
        {
            type: 'RichTextBlock',
            inlines: [
                {
                    type: 'TextRun',
                    text: 'Need to move app service devedsccpwu2-appservice1 to an app service plan on ASE. '
                }
            ]
        },
        {
            type: 'TextBlock',
            text: 'Troubleshooting:',
            wrap: true,
            weight: 'Bolder',
            color: 'Accent'
        },
        {
            type: 'RichTextBlock',
            inlines: [
                {
                    type: 'TextRun',
                    text: 'Have you already created the ASE and new app service plan?  What is the name of the destination ASE? When you want move an app service to an app service plan on an ASE you will have to create the ASE first then add the app service plan. From there you will have to create a new app service and clone the existing app service into the new app service. You can refer to Cloning an existing App to an App Service Environment using PowerShell or simply through the portal under the Clone App blade of the app service.'
                }
            ]
        },
        {
            type: 'TextBlock',
            text: 'Next Steps/Action Plan:',
            wrap: true,
            weight: 'Bolder',
            color: 'Accent'
        },
        {
            type: 'RichTextBlock',
            inlines: [
                {
                    type: 'TextRun',
                    text: 'For now kindly review information above and let me know if you have more questions and concerns on how to accomplish this.  If you have already tried and are receiving some sort of error please reply with an error message and screenshots of what you are running into. '
                }
            ]
        },
        {
            type: 'Container'
        },
        {
            type: 'TextBlock',
            text: 'Best Regards,',
            wrap: true
        },
        {
            type: 'Container'
        },
        {
            type: 'TextBlock',
            text: 'Gary D. Robinson',
            wrap: true,
            color: 'Accent',
            weight: 'Bolder'
        },
        {
            type: 'TextBlock',
            text: 'Support Engineer',
            wrap: true,
            weight: 'Bolder'
        },
        {
            type: 'TextBlock',
            text: 'Azure App Services',
            wrap: true
        },
        {
            type: 'TextBlock',
            text: '+1 (980) 776-2148',
            wrap: true
        },
        {
            type: 'TextBlock',
            text: 'garobins@microsoft.com',
            wrap: true
        },
        {
            type: 'TextBlock',
            text: 'Team Manager | Priscilla Morales | prmoral@microsoft.com ',
            wrap: true
        },
        {
            type: 'TextBlock',
            text: 'My normal working hours are Monday -Friday 8am-5pm US-Eastern Time.',
            wrap: true,
            size: 'Small',
            weight: 'Bolder'
        },
        {
            type: 'TextBlock',
            text: '\'Our mission is to empower every person and every organization on the planet to achieve more.\'',
            wrap: true,
            size: 'Small',
            color: 'Accent',
            weight: 'Bolder'
        }
    ],
    actions: [
        {
            type: 'Action.OpenUrl',
            title: 'Send',
            url: 'https://google.com'
        }
    ],
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.3'
});
