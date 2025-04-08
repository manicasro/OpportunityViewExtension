import styles from "../../view/ViewApplicationCustomizer.module.scss";

/**
 * Function to generate the title of the opportunity
 * @param title 
 * @returns 
 */
export function generateTitle(title: string): HTMLElement {
    const divElem = document.createElement('div');
    divElem.className = styles.opportunityTitleContainer;

    const titleParam = document.createElement('p');
    titleParam.className = styles.opportunityTitleParam;
    titleParam.innerHTML = 'Název zakázky';

    const titleValue = document.createElement('p');
    titleValue.className = styles.opportunityTitleValue;
    titleValue.innerHTML = title;

    divElem.appendChild(titleParam);
    divElem.appendChild(titleValue);

    return divElem;
}