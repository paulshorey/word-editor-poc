/* eslint-disable office-addins/no-context-sync-in-loop */
/* global console, window, setTimeout, Office, document, Word, require */
import { create } from "zustand";
import { TAGNAMES } from "@src/constants/constants";
import { ComponentTestData } from "@src/testdata/TestData";

/**
 * contentControl.id; context.document.contentControls.getById(id)
 */
export type id = number;
/**
 * contentControl.tag; context.document.contentControls.getByTag(tag);
 */
export type tag = string;

type outputOption = {
  id: string;
  title: string;
  condition: string;
};

export type dataElementType = {
  tag: tag;
  id: id;
  title?: string;
  outputOptions?: outputOption[];
};

export type componentsStateType = {
  items: dataElementType[];
  loadAll: () => Promise<Record<id, dataElementType>>;
  insertTag: (documentName: string) => dataElementType | undefined;
  //
  deleteId: (id: id) => Promise<dataElementType[]>;
  // deleteTags: (tag: tag) => Promise<dataElement[]>;
  //
  scrollToId: (id: id) => Promise<void>;
};

const componentsState = create((set, _get) => ({
  /**
   * All dataElements used in the template
   */
  items: [],

  loadAll: function () {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Read document
        const contentControls = context.document.contentControls.getByTag(TAGNAMES.component);
        context.load(contentControls, "items");
        await context.sync();
        // 2. Update state
        const all = [];
        for (let item of contentControls.items) {
          all.push({ id: item.id, tag: item.tag, title: item.title });
        }
        set({ items: all });
        resolve(null);
      });
    });
  },

  /**
   * Add a dataElement to the template, into the current cursor selection
   */
  insertTag: function (documentName: string) {
    return new Promise((resolve) => {
      Word.run(async (context) => {
        const contentRange = context.document.getSelection().getRange("Content");
        const contentControl = contentRange.insertContentControl();
        contentControl.set({
          tag: TAGNAMES.component,
          title: documentName.toUpperCase(),
        });
        await context.sync();

        switch (documentName) {
          case "helloworld_base64_plainText":
            {
              contentControl.load("insertFileFromBase64");
              await context.sync();
              await contentControl.context.sync();
              contentControl.insertFileFromBase64(ComponentTestData.helloworld_base64_plainText.data, "End");
            }
            break;

          case "helloworld_base64_dataURI":
            {
              contentControl.load("insertFileFromBase64");
              await context.sync();
              await contentControl.context.sync();
              contentControl.insertFileFromBase64(ComponentTestData.helloworld_base64_dataURI.data, "Replace");
            }
            break;

          case "helloworld_base64_xml":
            {
              contentControl.load("insertOoxml");
              await context.sync();
              await contentControl.context.sync();
              contentControl.insertOoxml(ComponentTestData.helloworld_base64_xml.data, "Replace");
            }
            break;

          case "helloworld_xml":
            {
              contentControl.load("insertOoxml");
              await contentControl.context.sync();
              await context.sync();
              contentControl.insertOoxml(ComponentTestData.helloworld_xml.data, "Replace");
            }
            break;

          case "comp_with_table":
            {
              contentControl.load("insertFileFromBase64");
              await contentControl.context.sync();
              await context.sync();
              contentControl.insertFileFromBase64(ComponentTestData.comp_with_table.data, "Replace");
            }
            break;

          case "comp_simple_word":
            {
              contentControl.load("insertFileFromBase64");
              await contentControl.context.sync();
              await context.sync();
              contentControl.insertFileFromBase64(ComponentTestData.comp_simple_word.data, "Replace");
            }
            break;

          default:
            Promise.reject("ERROR - Document does not exist");
            return;
        }
        await contentControl.context.sync();
        await context.sync();

        context.document.body.load();
        // await context.sync();
        context
          .sync()
          .then(() => {
            console.log("context.sync() SUCCESS");
          })
          .catch((error) => {
            console.error("context.sync() ERROR", error);
          });

        resolve(true);
        return;
      });

      setTimeout(async () => {
        Word.run(async (context) => {
          try {
            const after = context.document.getSelection().getRange("Content").getRange("After");
            await context.sync();
            after.load("insertHtml");
            // await context.sync();
            // after.insertHtml("<br />", "Start");
            await context.sync();
            after.insertHtml("<br />", "End");
            await context.sync();

            context.document.body.load();

            await context.sync();
          } catch (error) {
            console.error("try catch in componentState insertTag", error);
          }
        });
      }, 5000);

      // window.location.reload();
      // setTimeout(async () => {
      //   Word.run(async (context) => {
      //     const contentRange = context.document.getSelection().getRange("Start");
      //     await context.sync();
      //     const after = contentRange.getRange("After").insertContentControl();
      //     await context.sync();
      //     after.load("insertHtml");
      //     await context.sync();
      //     after.insertHtml("<br />", "Start");
      //     await context.sync();
      //     after.insertHtml("<br />", "End");
      //     await context.sync();
      //   });
      // }, 2000);
      // setTimeout(async () => {
      //   Word.run(async (context) => {
      //     const contentRange = context.document.getSelection().getRange("Start");
      //     await context.sync();
      //     const after = contentRange.getRange("After").insertContentControl();
      //     await context.sync();
      //     after.load("insertHtml");
      //     await context.sync();
      //     after.insertHtml("<br />", "End");
      //     await context.sync();
      //   });
      // }, 4000);
    });
  },

  deleteId: function (id: id): Promise<dataElementType[]> {
    // eslint-disable-next-line no-undef
    console.warn("dataElementsState.deleteTag()");
    return new Promise((resolve) => {
      Word.run(async (context) => {
        // 1. Delete from document
        const contentControl = context.document.contentControls.getById(id);
        await context.sync();
        contentControl.load("delete");
        await context.sync();
        contentControl.delete(false);
        await context.sync();
        // 2. Update state
        const all = await this.loadAll();
        resolve(all);
      });
    });
  },

  // deleteTags: function (tag: tag): Promise<dataElement[]> {
  //   console.warn("dataElementsState.deleteTag()");
  //   return new Promise((resolve) => {
  //     Word.run(async (context) => {
  //       // 1. Delete from document
  //       const contentControls = context.document.contentControls.getByTag(tag);
  //       context.load(contentControls, "items");
  //       await context.sync();
  //       for (let item of contentControls.items) {
  //         item.delete(false);
  //       }
  //       await context.sync();
  //       // 2. Update state
  //       const all = await this.loadAll();
  //       resolve(all);
  //     });
  //   });
  // },
}));

export default componentsState;
