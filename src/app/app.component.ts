import { Component } from '@angular/core';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  name = 'Angular 5.1';

  eventHandlers: Array<{
    workbook: Excel.Workbook;
    handler: OfficeExtension.EventHandlerResult<Excel.Workbook>;
  }> = [];

  constructor() {
    this.addEventHandler();
  }

  addEventHandler() {
    console.log('try adding');
    const onSelectionChanged = this.onSelectionChanged;
    const eventHandler = this.addEventHandler;
    this.tryCatch(() =>
      Excel.run(async context => {
        const workbook = context.workbook;
        // const handler = workbook.onSelectionChanged.add(onSelectionChanged);
        const handler = workbook.onSelectionChanged.add(async () => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'yellow';
          range.load('address');

          await context.sync();

          console.log(`New selection is ${range.address}`);
        });
        this.eventHandlers.push({ workbook, handler });

        await context.sync();

        console.log(
          'Event handler added',
          'Try changing the selection, and watch the console output.'
        );
      }).catch(function(error) {
        console.log('Error1: ' + error);
        if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
        console.log('call addEventHandler again');
        eventHandler();
      })
    );
  }

  async removeLastEventHandler() {
    const lastEventHandler = this.eventHandlers.pop();
    if (!lastEventHandler) {
      console.log('No event handlers added');
      return;
    }

    const workbook = lastEventHandler.workbook;
    this.tryCatch(() =>
      Excel.run(workbook, async context => {
        lastEventHandler.handler.remove();
        await context.sync();
      }).catch(function(error) {
        console.log('Error2: ' + error);
        if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
      })
    );
  }

  async removeAllEventHandlers() {
    if (this.eventHandlers.length === 0) {
      console.log('No event handlers added');
      return;
    }

    this.tryCatch(async () => {
      while (this.eventHandlers.length > 0) {
        const lastEventHandler = this.eventHandlers.pop();
        await Excel.run(lastEventHandler.workbook, async context => {
          lastEventHandler.handler.remove();
          await context.sync();
        }).catch(function(error) {
          console.log('Error3: ' + error);
          if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            console.error(error);
          }
        });
      }

      console.log('All event handlers removed');
    });
  }

  async onSelectionChanged() {
    this.tryCatch(() =>
      Excel.run(async context => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'yellow';
        range.load('address');

        await context.sync();

        console.log(`New selection is ${range.address}`);
      })
    );
  }

  /** Default helper for invoking an action and handling errors. */
  async tryCatch(callback: () => OfficeExtension.IPromise<any>) {
    try {
      await callback();
    } catch (error) {
      console.log(error);
      // OfficeHelpers.Utilities.log(error);
      this.addEventHandler();
    }
  }
}
