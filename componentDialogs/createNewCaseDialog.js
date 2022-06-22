const { WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { ConfirmPrompt, ChoicePrompt, DateTimePrompt, NumberPrompt, TextPrompt } = require('botbuilder-dialogs');
const { DialogSet, DialogTurnStatus } = require('botbuilder-dialogs');
// const { MessageFactory } = require('botbuilder');
// const { IndexCard } = require('./cards/IndexCard.js');
// const message = MessageFactory.attachment(IndexCard);

const CHOICE_PROMPT = 'CHOICE_PROMPT';
const CONFIRM_PROMPT = 'CONFIRM_PROMPT';
const TEXT_PROMPT = 'TEXT_PROMPT';
const NUMBER_PROMPT = 'NUMBER_PROMPT';
const DATETIME_PROMPT = 'DATETIME_PROMPT';
const WATERFALL_DIALOG = 'WATERFALL_DIAOG';
let endDialog = '';

class CreateNewCaseDialog extends ComponentDialog {
    constructor(conversationState, userState) {
        super('createNewCaseDialog');
        this.conversationState = conversationState;
        this.userState = userState;
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new NumberPrompt(NUMBER_PROMPT));
        this.addDialog(new DateTimePrompt(DATETIME_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [

            this.confirmCaseCreation.bind(this),
            this.getCaseID.bind(this),
            this.getDate.bind(this),
            this.getPrimaryContact.bind(this),
            this.getEmail.bind(this),
            this.getTitle.bind(this),
            this.getDescription.bind(this),
            this.getSeverity.bind(this),
            this.getScope.bind(this),
            this.confirmCaseInfo.bind(this),
            this.summary.bind(this)

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

    async confirmCaseCreation(step) {
        if (step._info.options.noOfParticipants) {
            step.values.noOfParticipants = step._info.options.noOfParticipants[0];
        }
        endDialog = false;
        return await step.prompt(CONFIRM_PROMPT, 'Are you sure you want to create a new case?', ['yes', 'no']);
    }

    async getCaseID(step) {
        if (step.result === true) {
            return await step.prompt(TEXT_PROMPT, 'What is the VDM CaseID?');
        }
    }

    async getDate(step) {
        step.values.caseID = step.result;
        return await step.prompt(DATETIME_PROMPT, 'On which day was the case created in VDM?');
    }

    async getPrimaryContact(step) {
        step.values.date = step.result;
        return await step.prompt(TEXT_PROMPT, 'Who is the primary contact for this case?');
    }

    async getEmail(step) {
        step.values.primaryContact = step.result;
        return await step.prompt(TEXT_PROMPT, 'What is their email?');
    }

    async getTitle(step) {
        step.values.email = step.result;
        return await step.prompt(TEXT_PROMPT, 'What is the case title?');
    }

    async getDescription(step) {
        step.values.title = step.result;
        return await step.prompt(TEXT_PROMPT, 'Provide a brief description of the issue.');
    }

    async getSeverity(step) {
        step.values.description = step.result;
        return await step.prompt(TEXT_PROMPT, 'What is the case Severity?');
    }

    async getScope(step) {
        step.values.serverity = step.result;
        return await step.prompt(CHOICE_PROMPT, 'What area of scope does this issue belong?', ['Create/Delete web app', 'VNET Integration', 'Hybrid Connection', 'Private Endpoint']);
    }

    async confirmCaseInfo(step) {
        var msg = ` You have entered the following values: 
        \n CaseID#: ${ step.values.caseID }
        \n Date: ${ step.values.date }
        \n Primary Contact: ${ step.values.primaryContact }
        \n Email: ${ step.values.email }
        \n Case Title: ${ step.values.title }
        \n Severity: ${ step.values.serverity }
        \n Scope: ${ step.values.scope }
        \n Description: ${ step.values.description }`;
        await step.context.sendActivity(msg);
        return await step.prompt(CONFIRM_PROMPT, 'Are you sure you all values are correct and you want to create this case?', ['yes', 'no']);
    }

    async summary(step) {
        if (step.result === true) {
            await step.context.sendActivity('Case was successfully created. You case id is: 13242454523');
            endDialog = true;
            return await step.endDialog();
        }
    }

    async isDialogComplete() {
        return endDialog;
    }
}

module.exports.CreateNewCaseDialog = CreateNewCaseDialog;
