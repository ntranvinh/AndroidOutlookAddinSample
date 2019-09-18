import * as React from "react";
import {ButtonType, PrimaryButton} from "office-ui-fabric-react";

export default class SuccessOwa extends React.Component {
    constructor(props, context) {
        super(props, context);
        this.state = { error: ""}
    }

    render() {
        let that = this
        return (
            <div className="left-owa">
                <p id="successfully-message">{this.props.message}</p>

                <div className="bottom-right-owa">
                    <PrimaryButton disabled={this.props.loading}
                                   buttonType={ButtonType.hero}
                                   onClick={function(){
                                       try{
                                           Office.context.ui.messageParent(JSON.stringify({ok: true}))
                                           console.log('user clicked ok')
                                       }
                                       catch(ex){
                                           that.setState({error: ex.message})
                                       }
                                   }}>OK</PrimaryButton>
                </div>
                <p>{this.state.error}</p>
            </div>
        );
    }
}