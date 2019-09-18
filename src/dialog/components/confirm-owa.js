import * as React from "react";
import {ButtonType, DefaultButton, PrimaryButton} from "office-ui-fabric-react";


export default class ConfirmOwa extends React.Component {
    constructor(props, context) {
        super(props, context);
        this.state = { error: ""}
    }
    render() {
        let that = this
        return (
            <div className="left-owa" >
                <p id="confirm-message">Confirm ?</p>

                <div className="bottom-right-owa">
                    <DefaultButton disabled={this.props.loading}
                                   className='ms-welcome__action' buttonType={ButtonType.hero}
                                   onClick={function(){
                                       try{
                                           Office.context.ui.messageParent(JSON.stringify({ok: false }))
                                           console.log('user clicked cancel')
                                       }
                                       catch(ex){
                                           that.setState({error: ex.message})
                                       }
                                   }}>Cancel</DefaultButton>
                    <PrimaryButton disabled={this.props.loading}
                                   className='ms-welcome__action' buttonType={ButtonType.hero}
                                   onClick={function(){
                                       try{
                                           Office.context.ui.messageParent(JSON.stringify({ok:true }))
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