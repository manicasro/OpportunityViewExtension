import { IOpportunity } from "../../../IOpportunity";
import { IConfig } from "../../config/Config";
import styles from "../../view/ViewApplicationCustomizer.module.scss";

/**
 * Function to generate buttons for Teams and SalesForce links
 * @param data 
 * @returns 
 */
export function generateButtons(data: IOpportunity, config: IConfig): HTMLElement {
    const divElem = document.createElement('div');
    divElem.className = styles.opportunityButtonContainer;

    // Generate Teams button if all required data is present
    if (data.sfaTeamId && data.sfaGenChannel && config?.tenantId && data.sfaTeamDone) {
        const teamsButton = createButton('Teams', styles.opportunityLinkButton, () => {
            const teamsUrl = `https://teams.microsoft.com/l/channel/${data.sfaGenChannel}/General?groupId=${data.sfaTeamId}&tenantId=${config.tenantId}`;
            window.open(teamsUrl, '_blank');
        });
        divElem.appendChild(teamsButton);
    }

    // Generate SalesForce button if opportunity or lead data is present
    const salesForceUrl = data.sfaOpportunityId
        ? `${config?.opportunityUrl ?? ''}${data.sfaOpportunityId}`
        : data.sfaLeadId
        ? `${config?.leadUrl ?? ''}${data.sfaLeadId}`
        : null;

    if (salesForceUrl) {
        const salesForceButton = createButton('SalesForce', styles.opportunityLinkButton, () => {
            window.open(salesForceUrl, '_blank');
        });
        divElem.appendChild(salesForceButton);
    }

    return divElem;
}

/**
 * Helper function to create a button element
 * @param text - Button text
 * @param className - CSS class for the button
 * @param onClick - Click event handler
 * @returns HTMLButtonElement
 */
function createButton(text: string, className: string, onClick: () => void): HTMLButtonElement {
    const button = document.createElement('button');
    button.className = className;
    button.innerHTML = text;
    button.addEventListener('click', onClick);
    return button;
}