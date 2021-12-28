import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'AceWithCascadingPropertyAdaptiveCardExtensionStrings';
import { IAceWithCascadingPropertyAdaptiveCardExtensionProps, IAceWithCascadingPropertyAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../AceWithCascadingPropertyAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IAceWithCascadingPropertyAdaptiveCardExtensionProps, IAceWithCascadingPropertyAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IBasicCardParameters {
    return {
      primaryText: strings.PrimaryText
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://www.bing.com'
      }
    };
  }
}
