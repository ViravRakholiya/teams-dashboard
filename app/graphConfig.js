// Add here the endpoints for MS Graph API services you would like to use.
const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages",
    graphUsersEndpoint: "https://graph.microsoft.com/v1.0/users",
    graphTeamEndpoint: "https://graph.microsoft.com/v1.0/users/{0}/licenseDetails?$select=servicePlans",
    graphUsersStatus: "https://graph.microsoft.com/v1.0/communications/getPresencesByUserId",
    graphUserEventsEndpoint:"https://graph.microsoft.com/v1.0/users/{0}/calendar/events?$orderby=start/datetime&$filter=start/datetime ge '{1}' and start/datetime le '{2}'",
    graphUserSearchEndPoint:'https://graph.microsoft.com/v1.0/users?$search="displayName:{0}" OR "mail:{0}"',
    graphGetGroupEndPoint: "https://graph.microsoft.com/v1.0/groups?$select=id,displayName",
    graphGetChannelEndPoint:"https://graph.microsoft.com/v1.0/teams/{0}/channels"
};
