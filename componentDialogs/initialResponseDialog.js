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

class InitialResponseDialog extends ComponentDialog {
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

            this.confirmInitialResponse.bind(this),
            this.getCaseID.bind(this),
            this.getDate.bind(this),
            this.getFirstName.bind(this),
            this.getEmail.bind(this),
            this.getTitle.bind(this),
            this.getIssueDescription.bind(this),
            this.getScope.bind(this),
            this.getTroubleshooting.bind(this),
            this.getActionPlan.bind(this),
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

    async confirmInitialResponse(step) {
        if (step._info.options.noOfParticipants) {
            step.values.noOfParticipants = step._info.options.noOfParticipants[0];
        }
        endDialog = false;
        return await step.prompt(CONFIRM_PROMPT, 'Are you sure you want to create a initial response?', ['yes', 'no']);
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
        step.values.date = step.result[0].value;
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

    async getScope(step) {
        step.values.description = step.result;
        return await step.prompt(CHOICE_PROMPT, {
            prompt: '"What are of Azure App App Services is the issue most closely associated with?',
            choices: [
                { value: 'Networking', action: { type: 'imBack', title: 'Networking', value: 'Networking' }, synonyms: ['Networking'] },
                { value: 'Management', action: { type: 'imBack', title: 'Management', value: 'Management' }, synonyms: ['Management'] },
                { value: 'Configuration', action: { type: 'imBack', title: 'Configuration', value: 'Configuration' }, synonyms: ['Configuration'] },
                { value: 'Authentication', action: { type: 'imBack', title: 'Authentication', value: 'Authentication' }, synonyms: ['Authentication'] },
                { value: 'Content Deployment', action: { type: 'imBack', title: 'Content Deployment', value: 'Content Deployment' }, synonyms: ['Content Deployment'] },
                { value: 'Development', action: { type: 'imBack', title: 'Development', value: 'Development' }, synonyms: ['Development'] }
            ]
        });
    }

    async getTroubleshooting(step) {
        step.values.scope = step.result;
        return await step.prompt(TEXT_PROMPT, 'What are some troubleshooting steps to start diagnosising this issue?');
    }

    async getActionPlan(step) {
        step.values.troubleshooting = step.result;
        return await step.prompt(TEXT_PROMPT, 'What are the next steps/action plan to resolve this issue?');
    }

    async preview(step) {
        step.values.action_plan = step.result;
        const EmailCard = CardFactory.adaptiveCard({
            type:'AdaptiveCard',
            $schema:'http://adaptivecards.io/schemas/adaptive-card.json',
            version:'1.3',
            body:[
               {
                  type:'RichTextBlock',
                  inlines:[
                     {
                        type:'TextRun',
                        text:`${ step.values.title } - Tracking# ${ step.values.caseID }`,
                        weight:'Bolder',
                        size:'Medium'
                     }
                  ],
                  id:'title',
                  separator:true
               },
               {
                  type:'ColumnSet',
                  columns:[
                     {
                        type:'Column',
                        width:'stretch',
                        items:[
                           {
                              type:'TextBlock',
                              text:`To: ${ step.values.email }`,
                              'isSubtle':true,
                              wrap:true,
                              size:'Small',
                              weight:'Bolder'
                           }
                        ],
                        id:'recipent'
                     },
                     {
                        type:'Column',
                        width:'stretch',
                        items:[
                           {
                              type:'TextBlock',
                              wrap:true,
                              size:'Small',
                              text:`Created: ${ step.values.date }`,
                              'isSubtle':true,
                              weight:'Bolder'
                           }
                        ],
                        id:'date'
                     }
                  ],
                  id:'recipent_date',
                  separator:true
               },
               {
                  type:'RichTextBlock',
                  inlines:[
                     {
                        type:'TextRun',
                        text:`Hi ${ step.values.firstName },\n \nMy name is Gary and I am from the App Services team. I can further assist you with troubleshooting the ${ step.values.scope.value } regarding your app service.  I have listed all of my backups and my managers contact information in my signature.\n\n`
                     }
                  ],
                  id:'greeting',
                  separator:true
               },
               {
                  type:'TextBlock',
                  text:'Issue Description:',
                  wrap:true,
                  weight:'Bolder',
                  id:'description'
               },
               {
                  type:'RichTextBlock',
                  inlines:[
                     {
                        type:'TextRun',
                        text:`${ step.values.description }`
                     }
                  ],
                  separator:true,
                  id:'description_text'
               },
               {
                  type:'TextBlock',
                  text:'Troubleshooting:',
                  wrap:true,
                  weight:'Bolder',
                  id:'troubleshooting'
               },
               {
                  type:'RichTextBlock',
                  inlines:[
                     {
                        type:'TextRun',
                        text:`${ step.values.troubleshooting }`
                     }
                  ],
                  id:'troubleshooting_text',
                  separator:true
               },
               {
                  type:'TextBlock',
                  text:'Next Steps/Action Plan:',
                  wrap:true,
                  weight:'Bolder',
                  id:'action_plan'
               },
               {
                  type:'RichTextBlock',
                  inlines:[
                     {
                        type:'TextRun',
                        text: `${ step.values.action_plan } \nFor now kindly review information above and let me know if you have more questions and concerns`
                     }
                  ],
                  separator:true,
                  id:'action_plan_text'
               },
               {
                  type:'Container',
                  id:'container_space'
               },
               {
                  type:'TextBlock',
                  text:'Best Regards,',
                  wrap:true,
                  id:'ending'
               },
               {
                  type:'Container',
                  id:'container_space_signature'
               },
               {
                  type:'Container',
                  items:[
                     {
                        type:'TextBlock',
                        text:'Gary D. Robinson',
                        wrap:true,
                        color:'Accent',
                        weight:'Bolder',
                        id:'name'
                     },
                     {
                        type:'TextBlock',
                        text:'Technical Support Engineer ',
                        wrap:true,
                        id:'job_title'
                     },
                     {
                        type:'TextBlock',
                        text:'Azure App Services ',
                        wrap:true,
                        size:'Small',
                        weight:'Bolder',
                        id:'support_area'
                     },
                     {
                        type:'TextBlock',
                        text:'Microsoft Customer Service and Support',
                        wrap:true,
                        size:'Small',
                        id:'department'
                     },
                     {
                        type:'TextBlock',
                        text:'Phone: +1 (980) 776-2148',
                        wrap:true,
                        size:'Small',
                        id:'phone'
                     },
                     {
                        type:'TextBlock',
                        text:'Email: garobins@microsoft.com',
                        wrap:true,
                        size:'Small',
                        id:'email'
                     },
                     {
                        type:'TextBlock',
                        text:'Team Manager | Priscilla Morales | prmoral@microsoft.com',
                        wrap:true,
                        size:'Small',
                        id:'manager'
                     },
                     {
                        type:'TextBlock',
                        text:'Backup | azurebu@microsoft.com',
                        wrap:true,
                        size:'Small',
                        id:'backup'
                     },
                     {
                        type:'TextBlock',
                        text:'My normal working hours are Mon-Friday 8am-5pm US-Eastern Time',
                        wrap:true,
                        size:'Small',
                        weight:'Bolder',
                        color:'Dark',
                        id:'working_hours'
                     },
                     {
                        type:'TextBlock',
                        text:'Our mission is to empower every person and every organization on the planet to achieve more.',
                        wrap:true,
                        size:'Small',
                        weight:'Bolder',
                        color:'Accent',
                        id:'mission'
                     },
                     {
                        type:'Image',
                        'url':'https://blogs.microsoft.com/wp-content/uploads/2012/08/8867.Microsoft_5F00_Logo_2D00_for_2D00_screen.jpg',
                        size:'Large',
                        id:'company_logo'
                     }
                  ],
                  id:'signature',
                  separator:true
               }
            ],
            id:'closing_email'
         });
        const message = MessageFactory.attachment(EmailCard);
        var msg = message;
        // await step.prompt(CONFIRM_PROMPT, 'Are you sure you all values are correct and you want to create this case?', ['yes', 'no']);
        return await step.context.sendActivity(msg);
    }

    async emailconfirmation(step) {
        if (step.result === true) {
            await step.context.sendActivity(`Initial Response for case ${ step.values.caseID } was successfully created.`);
            endDialog = true;
            return await step.endDialog();
        }
    }

    async isDialogComplete() {
        return endDialog;
    }
}

module.exports.InitialResponseDialog = InitialResponseDialog;
