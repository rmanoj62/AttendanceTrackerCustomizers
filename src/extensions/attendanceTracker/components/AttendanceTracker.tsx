import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

export default class AttendanceTrackerApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {

    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolder
    );

    return Promise.resolve();
  }

  private _renderPlaceHolder(): void {

    if (!this._topPlaceholder) {

      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top
        );

      if (!this._topPlaceholder) {
        console.log('Top placeholder not found');
        return;
      }

      this._topPlaceholder.domElement.innerHTML = `
        <div style="
            position:fixed;
            top:80px;
            right:20px;
            width:250px;
            background:white;
            border:1px solid #ccc;
            border-radius:8px;
            box-shadow:0 4px 12px rgba(0,0,0,.2);
            padding:15px;
            z-index:99999;
        ">
            <b>Attendance Tracker</b><br/><br/>
            Hello World
        </div>
      `;
    }
  }
}
