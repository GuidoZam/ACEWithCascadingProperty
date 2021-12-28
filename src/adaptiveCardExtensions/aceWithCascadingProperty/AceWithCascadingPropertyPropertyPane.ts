import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption, PropertyPaneDropdown, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'AceWithCascadingPropertyAdaptiveCardExtensionStrings';

export class AceWithCascadingPropertyPropertyPane {
  private parents: IPropertyPaneDropdownOption[] = [ { key: "A", text: "A" }, { key: "B", text: "B" }, { key: "C", text: "C" }];

  public getPropertyPaneConfiguration(children: IPropertyPaneDropdownOption[]): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('iconProperty', {
                  label: strings.IconPropertyFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true
                }),
                PropertyPaneDropdown("parent", {
                  label: "Parent",
                  options: this.parents
                }),
                PropertyPaneDropdown("child", {
                  label: "Child",
                  options: children
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
