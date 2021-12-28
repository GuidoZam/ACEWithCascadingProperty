import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { AceWithCascadingPropertyPropertyPane } from './AceWithCascadingPropertyPropertyPane';

export interface IAceWithCascadingPropertyAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  parent: string;
  child: string;
  children: IPropertyPaneDropdownOption[];
  enableAsync: boolean;
}

export interface IAceWithCascadingPropertyAdaptiveCardExtensionState {
  description: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'AceWithCascadingProperty_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'AceWithCascadingProperty_QUICK_VIEW';

export default class AceWithCascadingPropertyAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IAceWithCascadingPropertyAdaptiveCardExtensionProps,
  IAceWithCascadingPropertyAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: AceWithCascadingPropertyPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      description: this.properties.description
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo.svg');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'AceWithCascadingProperty-property-pane'*/
      './AceWithCascadingPropertyPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.AceWithCascadingPropertyPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration(this.properties.children);
  }

  protected setAsync = async (childValues: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
    console.log(this.properties.enableAsync);
    if(this.properties.enableAsync === true) {
      await this.delay(50);
    }
    return childValues;
  }

  protected delay(milliseconds: number) {
    return new Promise( resolve => setTimeout(resolve, milliseconds) );
  }

  protected onPropertyPaneFieldChanged = async (propertyPath: string, oldValue: any, newValue: any): Promise<void> => {
    if(propertyPath === "parent") {
      this.properties.child = undefined;
      const children = [...new Array(5)].map((value, index) => {const v = newValue + index; return { key: v, text: v };});
      switch(newValue) {
        case "A":
          this.properties.children = await this.setAsync(children);
        break;
        case "B":
          this.properties.children = await this.setAsync(children);
        break;
        case "C":
          this.properties.children = await this.setAsync(children);
        break;
      }
      console.log(this.properties.children);
    }
  }
}
