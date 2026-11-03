import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

export interface IAttendanceTrackerApplicationCustomizerProperties {
}

export default class AttendanceTrackerApplicationCustomizer
  extends BaseApplicationCustomizer<IAttendanceTrackerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {

    console.log('Attendance Tracker Loaded');

    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolder
    );

    this._renderPlaceHolder();

    return Promise.resolve();
  }

  private _renderPlaceHolder = (): void => {

    if (this._topPlaceholder) {
      return;
    }

    this._topPlaceholder =
      this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top
      );

    if (!this._topPlaceholder) {

      console.log('Top Placeholder Not Found');

      return;
    }

    this._topPlaceholder.domElement.innerHTML = `
      <div
        style="
          position:fixed;
          top:80px;
          right:20px;
          width:300px;
          background:#ffffff;
          border:1px solid #d1d1d1;
          border-radius:10px;
          padding:15px;
          box-shadow:0 4px 12px rgba(0,0,0,.15);
          z-index:99999;
          font-family:Segoe UI;
        "
      >
        <div
          style="
            display:flex;
            align-items:center;
            margin-bottom:10px;
          "
        >
          <span
            style="
              color:#107c10;
              font-size:18px;
              margin-right:8px;
            "
          >
            ⏰
          </span>

          <span
            style="
              font-size:16px;
              font-weight:600;
            "
          >
            Attendance Tracker
          </span>
        </div>

        <div style="margin-bottom:10px;">
          User:
          <b>${this.context.pageContext.user.displayName}</b>
        </div>

        <input
          id="attendanceComment"
          placeholder="Comments"
          style="
            width:100%;
            padding:8px;
            border:1px solid #ddd;
            border-radius:4px;
            box-sizing:border-box;
            margin-bottom:10px;
          "
        />

        <button
          id="attendanceBtn"
          style="
            width:100%;
            background:#d13438;
            color:white;
            border:none;
            padding:10px;
            border-radius:5px;
            cursor:pointer;
            font-weight:600;
          "
        >
          Sign In
        </button>
      </div>
    `;

    const button =
      document.getElementById(
        'attendanceBtn'
      ) as HTMLButtonElement;

    if (button) {

      button.addEventListener(
        'click',
        () => {

          const comment =
            (
              document.getElementById(
                'attendanceComment'
              ) as HTMLInputElement
            ).value;

          alert(
            'Sign In Clicked\n\nComment: ' +
            comment
          );
        }
      );
    }
  };

  @override
  protected onDispose(): void {

    if (this._topPlaceholder) {
      this._topPlaceholder.dispose();
    }

    super.onDispose();
  }
}