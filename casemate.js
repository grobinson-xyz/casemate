// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

class CaseMate extends ActivityHandler {
    constructor() {
        super();
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            const replyText = `Echo: ${ context.activity.text }`;
            await context.sendActivity(MessageFactory.text(replyText, replyText));
            // By calling next() you ensure that the next BotHandler is run.
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
        var reply = MessageFactory.suggestedActions(['Create New Case', 'Reply to Email', 'Create a Teams meeting'], 'Choose an option to get started');
        await turnContext.sendActivity(reply);
    }
}

module.exports.CaseMate = CaseMate;
