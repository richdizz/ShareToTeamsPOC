import { ContactGroupIcon, PersonIcon, TeamsMonochromeIcon } from "@fluentui/react-icons-northstar";
import { Dropdown, FormDropdown, Header } from "@fluentui/react-northstar";
import * as React from "react";

// component properties for the alert comonent
export interface PeopleChannelPickerProps { 
    selectionChangedHandler:any;
    accessToken:string;
    scope:number;
}

export interface PeopleChannelPickerState { 
    items:any[];
    selections:any[];
}
 
// ScopePicker component for displaying error messages
class PeopleChannelPicker extends React.Component<PeopleChannelPickerProps, PeopleChannelPickerState> {
    constructor(props:PeopleChannelPickerProps, context:PeopleChannelPickerState) {
        super(props, context);
        this.state = {
            items: [],
            selections: []
        }
    }

    componentDidMount() {
        // pre-load teams if it is a team picker
        if (this.props.scope == 3) {
            fetch("https://graph.microsoft.com/beta/me/joinedteams", {
                headers: { "Authorization": `Bearer ${this.props.accessToken}`}
            }).then((res: any) => {
                if (!res.ok) {
                    //TODO: error
                }
                else
                    return res.json();
            }).then((results) => {
                let items = [];
                results.value.forEach((i:any, n:number) => {
                    let item = {
                        id: i.id,
                        header: i.displayName,
                        image: "/images/nophoto.png",
                        content: i.description
                    };
                    items.push(item);

                    // async get photos
                    let ctx = item;
                    fetch(`https://graph.microsoft.com/beta/teams/${i.id}/photo/$value`, {
                        headers: { "Authorization": `Bearer ${this.props.accessToken}`}
                    }).then((res: any) => {
                        if (!res.ok) {
                            //TODO: error
                        }
                        else
                            return res.blob();
                    }).then((blob:any) => {
                        const url = window.URL || window.webkitURL;
                        const objectURL = url.createObjectURL(blob as Blob);
                        ctx.image = objectURL;
                        let list = this.state.items;
                        for (var i = 0; i < list.length; i++) {
                            if (list[i].id == ctx.id) {
                                list[i] = ctx;
                                break;
                            }
                        }
                        this.setState({items: list});
                    });
                });

                this.setState({items: items});
            });
        }
    }

    search(e:any) {
        if (this.props.scope != 3 && e.target.value.length > 1) {
            fetch(`https://graph.microsoft.com/v1.0/me/people?$search=${e.target.value}`, {
                headers: { "Authorization": `Bearer ${this.props.accessToken}`}
            }).then((res: any) => {
                if (!res.ok) {
                    //TODO: error
                }
                else
                    return res.json();
            }).then((results) => {
                let items = [];
                results.value.forEach((i:any, n:number) => {
                    if (i.personType.subclass == "OrganizationUser") {
                        let item = {
                            id: i.id,
                            header: i.displayName,
                            image: "/images/nophoto.png",
                            content: i.jobTitle
                        };
                        items.push(item);

                        // async get photos
                        let ctx = item;
                        fetch(`https://graph.microsoft.com/v1.0/users/${item.id}/photos/48x48/$value`, {
                            headers: { "Authorization": `Bearer ${this.props.accessToken}`}
                        }).then((res: any) => {
                            if (!res.ok) {
                                //TODO: error
                            }
                            else
                                return res.blob();
                        }).then((blob:any) => {
                            const url = window.URL || window.webkitURL;
                            const objectURL = url.createObjectURL(blob as Blob);
                            ctx.image = objectURL;
                            let list = this.state.items;
                            for (var i = 0; i < list.length; i++) {
                                if (list[i].id == ctx.id) {
                                    list[i] = ctx;
                                    break;
                                }
                            }
                            this.setState({items: list});
                        });
                    }
                });

                this.setState({items: items});
            });
        }
    }

    private selectionChanged(selectedItems: any) {
        if (this.props.scope != 2) {
            selectedItems = selectedItems[selectedItems.length - 1];
        }
        this.setState({selections: selectedItems});
        this.props.selectionChangedHandler(selectedItems);
    }
    
    // renders the component
    render() {
        return (
            <FormDropdown
                label="Destination"
                items={this.state.items}
                value={this.state.selections}
                placeholder="Select where message is sent"
                search
                multiple
                checkable
                fluid
                onKeyDown={this.search.bind(this)}
                noResultsMessage="No results found"
                onChange={(_evt, ctrl) => this.selectionChanged(ctrl.value)}
            />
        );
    };
}
 
export default PeopleChannelPicker;