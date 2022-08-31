import { PublicClientApplication } from "@azure/msal-browser";
import * as React from "react";
import { Button, FormDropdown, FormInput, FormTextArea, Header, Loader, TextArea } from "@fluentui/react-northstar";
import * as AdaptiveCards from "adaptivecards";
import * as ReactDOM from "react-dom";
import ScopePicker from "./ScopePicker";
import PeopleChannelPicker from "./PeopleChannelPicker";

// component properties
export interface AppProps {
}

// component state
export interface AppState {
    waiting: boolean;
    error: string;
    ready: boolean;
    card: AdaptiveCards.AdaptiveCard;
    selectedScope?: number;
    accessToken?:string;
    cardJson?:any;
    groupName:string;
    message:string;
    recipients?:any;
    channels:any[];
    selectedChannel?:any;
    submit: boolean;
}

// App component
export default class App extends React.Component<AppProps, AppState> {
    myMSALObj: PublicClientApplication;
    urlParams = new URLSearchParams(window.location.search);
    constructor(props:AppProps, context:AppState) {
        super(props, context);
        this.state = {
            waiting: false,
            error: null,
            ready: false,
            card: null,
            message: "",
            groupName: "",
            channels: [],
            submit: false
        };

        var client_id = "02b0f240-abc2-4db4-9851-80d40f9a3abd";
        this.myMSALObj = new PublicClientApplication({
            auth: {
                clientId: client_id,
                authority: "https://login.microsoftonline.com/richdizz.com/"
            },
            cache: {
                cacheLocation: "sessionStorage",
                storeAuthStateInCookie: false
            }
        });
    };

    componentDidMount = async () => {
        // HACK...get client_id from url
        await this.myMSALObj.handleRedirectPromise();
        const accounts = this.myMSALObj.getAllAccounts();

        // listen for adaptive card payload from client
        let ctx = this;
        window.addEventListener("message", async (event) => {
            if (event.origin == "http://localhost:3474") {
                console.log("Card recieved from host");
                // save the payload into session storage and then ensure user logged in
                window.sessionStorage.setItem("card", event.data);
                const accounts = this.myMSALObj.getAllAccounts();
                if (accounts.length === 0) {
                    // No user signed in
                    var t = await this.myMSALObj.acquireTokenRedirect({
                        scopes: [
                            "Channel.ReadBasic.All",
                            "ChannelMessage.Send",
                            "Chat.Create",
                            "ChatMessage.Send",
                            "Team.ReadBasic.All",
                            "User.Read",
                            "User.ReadBasic.All",
                            "People.Read"
                        ],
                        redirectStartPage: window.location.href
                    });
                }
            }
        });

        // if the user is signed in, acquire the token
        if (accounts.length != 0) {
            var resp = await this.myMSALObj.acquireTokenSilent({
                scopes: [
                    "Channel.ReadBasic.All",
                    "ChannelMessage.Send",
                    "Chat.Create",
                    "ChatMessage.Send",
                    "Team.ReadBasic.All",
                    "User.Read",
                    "User.ReadBasic.All",
                    "People.Read"
                ],
                account: accounts[0]
            });
            if (resp.accessToken) {
                var cardPayload = window.sessionStorage.getItem("card");
                var cardJson = JSON.parse(cardPayload);
                let card = new AdaptiveCards.AdaptiveCard();
                card.onExecuteAction = function(action) { alert("Ow!"); }
                card.parse(cardJson);

                this.setState({ready: true, card: card, cardJson: cardJson, accessToken: resp.accessToken});
            }
        }
    }

    onScopeSelected = (scope:number) => {
        this.setState({selectedScope: scope});
    };

    sendMessage = () => {
        this.setState({submit: true});
        let payload = {
            "subject": null,
            "body": {
                "contentType": "html",
                "content": `${this.state.message} <attachment id=\"74d20c7f34aa4a7fb74e2b30004247c8\"></attachment>`
            },
            "attachments": [
                {
                    "id": "74d20c7f34aa4a7fb74e2b30004247c8",
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "contentUrl": null,
                    "content": JSON.stringify(this.state.cardJson),
                    "name": null,
                    "thumbnailUrl": null
                }
            ]
        };

        if (this.state.selectedScope == 1 || this.state.selectedScope == 2) {
            // build the members section
            let members = [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": "https://graph.microsoft.com/beta/users('ccaf48a2-0fe3-4e2d-8913-829c0f1dcfac')"
                }
            ];
            if (this.state.selectedScope == 1) {
                members.push({
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": `https://graph.microsoft.com/beta/users('${this.state.recipients.id}')`
                });
            }
            else {
                for (var i = 0; i < this.state.recipients.length; i++) {
                    members.push({
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": `https://graph.microsoft.com/beta/users('${this.state.recipients[i].id}')`
                    });
                }
            }

            let scopeText = (this.state.selectedScope == 1) ? "oneOnOne" : "group";
            let chatPayload = {
                "chatType": scopeText,
                "topic": null,
                "members": members
            };
            if (this.state.selectedScope == 2) {
                chatPayload.topic = this.state.groupName;
            }


            fetch("https://graph.microsoft.com/beta/chats", {
                method: "POST",
                body: JSON.stringify(chatPayload),
                headers: {
                    "accept": "application/json",
                    "content-type": "application/json",
                    "Authorization": `Bearer ${this.state.accessToken}`
                }
            }).then((res:any) => {
                console.log(res);
                if (res.ok)
                    return res.json();
            }).then((res:any) => {
                // now send the message
                fetch(`https://graph.microsoft.com/beta/chats/${res.id}/messages`, {
                    method: "POST",
                    body: JSON.stringify(payload),
                    headers: {
                        "accept": "application/json",
                        "content-type": "application/json",
                        "Authorization": `Bearer ${this.state.accessToken}`
                    }
                }).then((r:any) => {
                    if (r.ok)
                        return r.json();
                }).then((r:any) => {
                    window.close();
                });
            });
        }
        else if (this.state.selectedScope == 3) {
            // send to team channel
            console.log(this.state.selectedChannel);
            fetch(`https://graph.microsoft.com/beta/teams/${this.state.recipients.id}/channels/${this.state.selectedChannel.id}/messages`, {
                method: "POST",
                body: JSON.stringify(payload),
                headers: {
                    "accept": "application/json",
                    "content-type": "application/json",
                    "Authorization": `Bearer ${this.state.accessToken}`
                }
            }).then((r:any) => {
                if (r.ok)
                    return r.json();
            }).then((r:any) => {
                window.close();
            });
        }
    };

    onSendToChanged(to:any) {
        // get channels for this team if team selected
        if (this.state.selectedScope == 3) {
            fetch(`https://graph.microsoft.com/beta/teams/${to.id}/channels`, {
                headers: { "Authorization": `Bearer ${this.state.accessToken}`}
            }).then((res: any) => {
                if (res.ok) {
                    return res.json();
                }
            }).then((res:any) => {
                let channels = [];
                for (var i = 0; i < res.value.length; i++) {
                    channels.push({
                        id: res.value[i].id,
                        header: res.value[i].displayName,
                        content: res.value[i].description
                    });
                }
                this.setState({channels: channels});
            });
        }

        // save the state
        this.setState({recipients: to});
    };

    // renders the component
    render() {
        var payload = (<></>);
        if (!this.state.ready) {
            payload = (<Loader label="Loading..." size="large" />);
        }
        else if (!this.state.selectedScope) {
            payload = (<ScopePicker scopeSelectedHandler={this.onScopeSelected.bind(this)} />);
        }
        else {
            let scope = (this.state.selectedScope == 1) ? "Individual" : ((this.state.selectedScope == 2) ? "Group" : "Team Channel");
            let groupName = (this.state.selectedScope == 2) ? (<FormInput fluid label="Group name" placeholder="Type a name for this group..." value={this.state.groupName} onChange={(e:any) => this.setState({groupName: e.target.value})} />) : (<></>);
            let channelSelector = (this.state.selectedScope == 3) ? (<FormDropdown label="Channel" fluid items={this.state.channels} value={this.state.selectedChannel} onChange={(_evt, ctrl) => this.setState({selectedChannel: ctrl.value})}  />) : (<></>);
            let cardPreview = this.state.card.render();
            payload = (
                <div>
                    <Header as="h3" content={"Share to " + scope} />
                    
                    <PeopleChannelPicker accessToken={this.state.accessToken} scope={this.state.selectedScope} selectionChangedHandler={this.onSendToChanged.bind(this)} />
                    {groupName}
                    {channelSelector}
                    <FormTextArea label="Optional message" fluid placeholder="Type an optional message..." value={this.state.message} onChange={(e:any) => this.setState({message: e.target.value})} />
                    <Header as="h6" content="Card preview" style={{paddingTop: "10px"}} />
                    <div style={{border: "1px solid #ccc", padding: "10px"}} dangerouslySetInnerHTML={{ __html: cardPreview?.innerHTML }}></div>
                    <div style={{position: "fixed", right: "20px", bottom: "20px"}}>
                        <Button onClick={() => window.close()} content="Cancel" style={{marginRight: "10px"}} />
                        <Button onClick={this.sendMessage.bind(this)} content="Send" primary loading={this.state.submit} />
                    </div>
                </div>
            );
        }

        return (
        <div style={{paddingLeft: "20px", paddingRight: "20px", paddingTop: "20px"}}>
            {payload}
        </div>
        );
    };
}