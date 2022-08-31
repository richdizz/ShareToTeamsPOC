import { ContactGroupIcon, PersonIcon, TeamsMonochromeIcon } from "@fluentui/react-icons-northstar";
import { Header } from "@fluentui/react-northstar";
import * as React from "react";

// component properties for the alert comonent
export interface ScopePickerProps { 
    scopeSelectedHandler:any;
}
 
// ScopePicker component for displaying error messages
class ScopePicker extends React.Component<ScopePickerProps> {
    // renders the component
    render() {
        return (
            <div>
                <Header as="h3" content="Share to..." />
                <div className="scopeItem" onClick={() => this.props.scopeSelectedHandler(1)}>
                    <div className="scopeIcon"><PersonIcon size="largest" /></div>
                    <div className="scopeText">Individual</div>
                </div>
                <div className="scopeItem" onClick={() => this.props.scopeSelectedHandler(2)}>
                    <div className="scopeIcon"><ContactGroupIcon size="largest" /></div>
                    <div className="scopeText">Group</div>
                </div>
                <div className="scopeItem" onClick={() => this.props.scopeSelectedHandler(3)}>
                    <div className="scopeIcon"><TeamsMonochromeIcon size="largest" /></div>
                    <div className="scopeText">Team Channel</div>
                </div>
            </div>
        );
    };
}
 
export default ScopePicker;