const { CardFactory } = require('botbuilder');

exports.IndexCard = CardFactory.adaptiveCard({
    type: 'AdaptiveCard',
    body: [
        {
            type: 'TextBlock',
            text: 'Hello Gary, chose an option to get started.',
            wrap: true,
            size: 'Large'
        },
        {
            type: 'Container',
            items: [
                {
                    type: 'ColumnSet',
                    columns: [
                        {
                            type: 'Column',
                            width: 'stretch',
                            items: [
                                {
                                    type: 'ActionSet',
                                    actions: [
                                        {
                                            type: 'Action.Submit',
                                            title: 'Create New Case'
                                        }
                                    ],
                                    $data: 'Create New Case'
                                }
                            ]
                        },
                        {
                            type: 'Column',
                            width: 'stretch',
                            items: [
                                {
                                    type: 'ActionSet',
                                    actions: [
                                        {
                                            type: 'Action.Submit',
                                            title: 'Open my Notes'
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            type: 'Column',
                            width: 'stretch',
                            items: [
                                {
                                    type: 'ActionSet',
                                    actions: [
                                        {
                                            type: 'Action.Submit',
                                            title: 'Create a repo'
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        },
        {
            type: 'Container',
            items: [
                {
                    type: 'TextBlock',
                    size: 'Medium',
                    weight: 'Bolder',
                    text: 'Casebuddy Dashboard'
                },
                {
                    type: 'TextBlock',
                    text: 'June 1, 2022 - June 11, 2022',
                    wrap: true
                }
            ]
        },
        {
            type: 'ColumnSet',
            columns: [
                {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                        {
                            type: 'TextBlock',
                            wrap: true,
                            text: '10',
                            size: 'ExtraLarge'
                        }
                    ]
                },
                {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                        {
                            type: 'TextBlock',
                            text: '3',
                            wrap: true,
                            size: 'ExtraLarge'
                        }
                    ]
                },
                {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                        {
                            type: 'TextBlock',
                            text: '4.3',
                            wrap: true,
                            size: 'ExtraLarge'
                        }
                    ]
                }
            ]
        },
        {
            type: 'ColumnSet',
            columns: [
                {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                        {
                            type: 'TextBlock',
                            text: 'Open Case',
                            wrap: true
                        }
                    ]
                },
                {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                        {
                            type: 'TextBlock',
                            text: 'Closed ',
                            wrap: true
                        }
                    ]
                },
                {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                        {
                            type: 'TextBlock',
                            text: 'CSAT',
                            wrap: true
                        }
                    ]
                }
            ]
        },
        {
            type: 'TextBlock',
            text: 'Cases',
            wrap: true,
            size: 'Large'
        },
        {
            type: 'Container',
            items: [
                {
                    type: 'TextBlock',
                    text: '342432420349243',
                    wrap: true,
                    weight: 'Bolder'
                },
                {
                    type: 'ColumnSet',
                    columns: [
                        {
                            type: 'Column',
                            width: 'stretch',
                            items: [
                                {
                                    type: 'ActionSet',
                                    actions: [
                                        {
                                            type: 'Action.Submit',
                                            title: 'Hybrid connection unable to resolve to correct IP address'
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        },
        {
            type: 'Container',
            items: [
                {
                    type: 'TextBlock',
                    text: '534435234553454',
                    wrap: true,
                    weight: 'Bolder'
                }
            ]
        },
        {
            type: 'ColumnSet',
            columns: [
                {
                    type: 'Column',
                    width: 'stretch',
                    items: [
                        {
                            type: 'ActionSet',
                            actions: [
                                {
                                    type: 'Action.Submit',
                                    title: 'Private Endpoint fails to connect to On prem resources'
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ],
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.3'
});
