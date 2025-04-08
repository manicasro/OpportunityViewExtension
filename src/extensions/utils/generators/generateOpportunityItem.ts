import styles from "../../view/ViewApplicationCustomizer.module.scss";

/**
 * Function to generate an opportunity item (name + value pair)
 * @param parameterName 
 * @param parameterValue 
 * @returns 
 */
export function generateOpportunityItem(parameterName: string, parameterValue: string): HTMLElement {
    if (!parameterValue) {
      parameterValue = 'N/A';
    }
    
    const divElem = document.createElement('div');
    divElem.className = styles.opportunityItemView;

    const parameterNamePar = document.createElement('p');
    parameterNamePar.className = styles.opportunityItemParamName;
    parameterNamePar.innerHTML = parameterName;

    const parameterValuePar = document.createElement('p');
    parameterValuePar.className = styles.opportunityItemParamValue;
    parameterValuePar.innerHTML = parameterValue;

    divElem.appendChild(parameterNamePar);
    divElem.appendChild(parameterValuePar);

    return divElem;
  }