export default class Person {
    Text: string;
    TertiaryText: string;
    PictureURL: string;
    SecondaryText: string;
    Date: Date;
    constructor(webUrl: string, item: any, textField: string, secondaryTextField: string, tertiaryTextField: string, dateFieldString: string) {
        this.Text = item[textField];
        this.PictureURL = item['WorkEmail']!=null ? `${webUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${item['WorkEmail']}`: '#';
        this.SecondaryText = item[secondaryTextField];
        this.TertiaryText = item[tertiaryTextField];
        this.Date = item[dateFieldString];
    }
}