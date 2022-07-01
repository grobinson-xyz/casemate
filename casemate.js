// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');
// const { IndexCard } = require('./cards/IndexCard.js');
// const message = MessageFactory.attachment(IndexCard);

const { CreateNewCaseDialog } = require('./componentDialogs/createNewCaseDialog');
const { ClosingEmailDialog } = require('./componentDialogs/closingEmailDialog');

class CaseMate extends ActivityHandler {
    constructor(conversationState, userState) {
        super();
        this.conversationState = conversationState;
        this.userState = userState;
        this.dialogState = conversationState.createProperty('dialogState');
        this.createNewCaseDialog = new CreateNewCaseDialog(this.conversationState, this.userState);
        this.closingEmailDialog = new ClosingEmailDialog(this.conversationState, this.userState);
        this.previousIntent = this.conversationState.createProperty('previousIntent');
        this.conversationData = this.conversationState.createProperty('conversationData');
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            await this.dispatchToIntentAsync(context);

            await next();
        });

        this.onDialog(async (context, next) => {
            await this.conversationState.saveChanges(context, false);
            await this.userState.saveChanges(context, false);

            await next();
        });

        this.onMembersAdded(async (context, next) => {
            await this.sendWelcomeMessage(context);
            // const membersAdded = context.activity.membersAdded;
            // const welcomeText = 'Hello and welcome!';
            // for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
            //     if (membersAdded[cnt].id !== context.activity.recipient.id) {
            //         await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
            //     }
            // }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    async sendWelcomeMessage(turnContext) {
        const { activity } = turnContext;
        for (const idx in activity.membersAdded) {
            if (activity.membersAdded[idx].id !== turnContext.activity.recipient.id) {
                const welcomeMessage = `Welcome to Casemate! ${ activity.membersAdded[idx].name }.`;
                await turnContext.sendActivity(welcomeMessage);
                await this.sendSuggestedActions(turnContext);
            }
        }
    }

    async sendSuggestedActions(turnContext) {
        // var reply = await turnContext.sendActivity(message);
        var reply = MessageFactory.suggestedActions(['Create New Case', 'Closing Email'], 'Chose an Option to get started.');
        await turnContext.sendActivity(reply);
    }

    async dispatchToIntentAsync(context) {
        let currentIntent = '';
        const previousIntent = await this.previousIntent.get(context, {});
        const conversationData = await this.conversationData.get(context, {});
        if (previousIntent.intentName && conversationData.endDialog === false) {
            currentIntent = previousIntent.intentName;
        } else if (previousIntent.intentName && conversationData.endDialog === true) {
            currentIntent = context.activity.text;
        } else {
            currentIntent = context.activity.text;
            await this.previousIntent.set(context, { intentName: context.activity.text });
        }
        switch (currentIntent) {
        case 'Create New Case':
            console.log('Inside Create New Case dialog');
            await this.conversationData.set(context, { endDialog: false });
            await this.createNewCaseDialog.run(context, this.dialogState);
            conversationData.endDialog = await this.createNewCaseDialog.isDialogComplete();
            if (conversationData.endDialog) {
                await this.sendSuggestedActions(context);
            }
            break;
        case 'Closing Email':
            console.log('Inside Closing Email dialog');
            await this.conversationData.set(context, { endDialog: false });
            await this.closingEmailDialog.run(context, this.dialogState);
            conversationData.endDialog = await this.closingEmailDialog.isDialogComplete();
            if (conversationData.endDialog) {
                await this.sendSuggestedActions(context);
            }
            break;
        default:
            console.log('Did not match About me case');
            break;
        }
    }
}

module.exports.CaseMate = CaseMate;
