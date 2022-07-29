/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

const docs = [
  {
    thumbnail: "../../assets/doc-1.png",
    title: "Document 1",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/721ae21344f7ca71c3ed184c87aa1146/raw/027085a8d76a6818e2e4a68de2653fed2250a2eb/olympta",
  },
  {
    thumbnail: "../../assets/doc-2.png",
    title: "Document 1",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/8aa719d5e61bde40d249a3e49c691591/raw/4e4c54c1c5ec3fa405dea085f9ae680de06975a0/document%25202",
  },
  {
    thumbnail: "../../assets/doc-3.png",
    title: "Document 1",
    description: "This is a document",
    url: "https://gist.githubusercontent.com/chibchombiano26/8aa719d5e61bde40d249a3e49c691591/raw/4e4c54c1c5ec3fa405dea085f9ae680de06975a0/document%25202",
  },
];

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("send-to-beacon").addEventListener("click", () => {
      Office.context.document.getFileAsync(Office.FileType.Pdf, (result) => {
        // @ts-ignore
        if (result.status == "succeeded") {
          const myFile = result.value;
          myFile.getSliceAsync(0, (slice) => {
            // @ts-ignore
            if (slice.status == "succeeded") {
              const fileContent = slice.value.data;
              // byte array to blob
              const blob = new Blob([fileContent], { type: "application/pdf" });
              // download blob
              const url = URL.createObjectURL(blob);
              const a = document.createElement("a");
              a.href = url;
              a.download = "myFile.pdf";
              document.body.appendChild(a);
              a.click();
              document.body.removeChild(a);
            }
          });

          myFile.closeAsync();
        } else {
          // app.showNotification("Error:", result.error.message);
        }
      });
    });

    generateThumbnails();
  }
});

function generateThumbnails() {
  docs.forEach((doc) => {
    const node = document.createElement("div");
    const img = document.createElement("img");
    img.src = doc.thumbnail;
    img.onclick = () => {
      uploadFile(doc.url);
    };
    node.appendChild(img);
    document.getElementById("thumbnails").appendChild(node);
  });
}

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}

const uploadFile = async (url: string) => {
  const blob = await fetch(url);
  const text = await blob.text();
  PowerPoint.createPresentation(text);
};
