const { WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { ConfirmPrompt, ChoicePrompt, DateTimePrompt, NumberPrompt, TextPrompt } = require('botbuilder-dialogs');
const { DialogSet, DialogTurnStatus } = require('botbuilder-dialogs');
const { CardFactory } = require('botbuilder');
const { MessageFactory } = require('botbuilder');
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const CONFIRM_PROMPT = 'CONFIRM_PROMPT';
const TEXT_PROMPT = 'TEXT_PROMPT';
const NUMBER_PROMPT = 'NUMBER_PROMPT';
const DATETIME_PROMPT = 'DATETIME_PROMPT';
const WATERFALL_DIALOG = 'WATERFALL_DIAOG';
let endDialog = '';

class ClosingEmailDialog extends ComponentDialog {
    constructor(conversationState, userState) {
        super('closingEmailDialog');
        this.conversationState = conversationState;
        this.userState = userState;
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new NumberPrompt(NUMBER_PROMPT));
        this.addDialog(new DateTimePrompt(DATETIME_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [

            this.confirmClosingEmail.bind(this),
            this.getCaseID.bind(this),
            this.getDate.bind(this),
            this.getFirstName.bind(this),
            this.getEmail.bind(this),
            this.getTitle.bind(this),
            this.getIssueDescription.bind(this),
            this.getAnalysis.bind(this),
            // this.getScope.bind(this),
            this.preview.bind(this),
            this.emailconfirmation.bind(this)

        ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    async run(turnContext, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);
        const dialogContext = await dialogSet.createContext(turnContext);

        const result = await dialogContext.continueDialog();
        if (result.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async confirmClosingEmail(step) {
        if (step._info.options.noOfParticipants) {
            step.values.noOfParticipants = step._info.options.noOfParticipants[0];
        }
        endDialog = false;
        return await step.prompt(CONFIRM_PROMPT, 'Are you sure you want to send a closing email?', ['yes', 'no']);
    }

    async getCaseID(step) {
        if (step.result === true) {
            return await step.prompt(TEXT_PROMPT, 'What is the VDM CaseID?');
        }
    }

    async getDate(step) {
        step.values.caseID = step.result;
        return await step.prompt(DATETIME_PROMPT, "What is today's date?");
    }

    async getFirstName(step) {
        step.values.date = step.result;
        return await step.prompt(TEXT_PROMPT, "What is the recipient's first name?");
    }

    async getEmail(step) {
        step.values.firstName = step.result;
        return await step.prompt(TEXT_PROMPT, 'What is their email?');
    }

    async getTitle(step) {
        step.values.email = step.result;
        return await step.prompt(TEXT_PROMPT, 'What is the case title?');
    }

    async getIssueDescription(step) {
        step.values.title = step.result;
        return await step.prompt(TEXT_PROMPT, 'Provide a brief description of the issue.');
    }

    async getAnalysis(step) {
        step.values.description = step.result;
        return await step.prompt(TEXT_PROMPT, 'What was the root case anaylsis or guidance provided?');
    }

    // async getScope(step) {
    //     step.values.serverity = step.result;
    //     return await step.prompt(CHOICE_PROMPT, 'What area of scope does this issue belong?', ['Create/Delete web app', 'VNET Integration', 'Hybrid Connection', 'Private Endpoint']);
    // }

    async preview(step) {
        step.values.rca = step.result;
        const EmailCard = CardFactory.adaptiveCard({
            type: 'AdaptiveCard',
            body: [
                {
                    type: 'TextBlock',
                    size: 'Medium',
                    weight: 'Bolder',
                    text: `${ step.values.title } - Tracking# ${ step.values.caseID }`
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
                                    text: `To: ${ step.values.email }`,
                                    wrap: true
                                },
                                {
                                    type: 'TextBlock',
                                    spacing: 'None',
                                    text: 'From: garobins@microsoftsupport.com',
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
                    text: `Hi ${ step.values.firstName },\n \nMy name is Gary and I am from the App Services team. I can further assist you with troubleshooting your migrating your app service.  I have listed all of my backups and my managers contact information in my signature.\n\n`,
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
                            text: `${ step.values.description }`
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
                            text: `${ step.values.rca }`
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
        const message = MessageFactory.attachment(EmailCard);
        var msg = message;
        // await step.prompt(CONFIRM_PROMPT, 'Are you sure you all values are correct and you want to create this case?', ['yes', 'no']);
        return await step.context.sendActivity(msg);
    }

    async emailconfirmation(step) {
        if (step.result === true) {
            await step.context.sendActivity(`Case ${ step.values.caseID } was successfully closed.`);
            endDialog = true;
            return await step.endDialog();
        }
    }

    async isDialogComplete() {
        return endDialog;
    }
}

module.exports.ClosingEmailDialog = ClosingEmailDialog;
