const duplicateWorkbook = () => {
  Office.context.document.getFileAsync(Office.FileType.Compressed, async ({ value: fileHandle }) => {
    const buf = new Uint8Array(fileHandle.size);
    let offset = 0;

    for (let i = 0; i < fileHandle.sliceCount; i++) {
      await new Promise<void>((resolve) => {
        fileHandle.getSliceAsync(i, ({ value: slice }) => {
          buf.set(slice.data as Array<number>, offset);
          offset += slice.size;
          resolve();
        });
      });
    }

    const fileBinary = new Blob([buf]);

    const workbookB64 = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        resolve(reader.result as string);
      };
      reader.onerror = reject;
      reader.readAsDataURL(fileBinary);
    });

    const dataOffset = workbookB64.indexOf("base64,");
    const b64Offset = workbookB64.substring(dataOffset + "base64,".length);
    fileHandle.closeAsync();

    await Excel.createWorkbook(b64Offset);
  });
};

Office.onReady(() => {
  document.getElementById("duplicate-btn").onclick = duplicateWorkbook;

  Excel.run(async (context) => {
    context.workbook.worksheets.onCalculated.add((ev) => {
      console.info("onCalculated", ev);
    });
  });
});
