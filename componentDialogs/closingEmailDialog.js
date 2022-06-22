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
        var msg = ` Email preview:
        \n 
        \n Subject: ${ step.values.title } - Tracking# ${ step.values.caseID }
        \n Hi ${ step.values.firstName }
        \n
        \n
        \n Thanks for your confirmation and partnering with me on your case. Per your confirmation I'll be closing this
        case. Please keep in mind that when your case is closed, we can reopen it at any time.
        \n
        \n Below, you can see the summary of your case for your records. 
        \n
        \n Issue: 
        \n${ step.values.description }
        \n 
        \n
        \n Outcome:
        \n${ step.values.rca }
        \n
        \n
        \n
        \n I appreciate your time working together to resolve this issue. If you have any feedback or concerns on the handling of your case, please feel free to reach out to my Priscilla Morales at prmoral@microsoft.com, or listed below in my signature.

        \n Your feedback is important to us. After this interaction, you will receive a separate closure email with an opportunity to tell us about your experience.
        \n
        \n We wish you the best with your future development! 
        \n Thanks again,
        \n`;
        await step.context.sendActivity(msg);
        return await step.prompt(CONFIRM_PROMPT, 'Are you sure you all values are correct and you want to create this case?', ['yes', 'no']);
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
