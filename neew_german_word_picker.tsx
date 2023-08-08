function main(workbook: ExcelScript.Workbook) {
    const selectedSheet = workbook.getActiveWorksheet();

    type wordPicker = {
        email: string;
        hasDoneDuty: string | number | boolean;
    }

    const range = selectedSheet.getRange("A2:B20");
    const nonEmptyValues = range.getValues().filter(row => row[0] != "");

    const allWordPickers: Array<wordPicker> = nonEmptyValues.map((row) => {
        return {
            email: row[0].toString(),
            hasDoneDuty: row[1]
        }
    })

    function getAvailableWordPickers(): Array<wordPicker> {
      const availableWordPickers: Array<wordPicker> = []
      allWordPickers.forEach(wordPicker => {
        if (!wordPicker.hasDoneDuty) {
          availableWordPickers.push(wordPicker);
        }
      })

      if (availableWordPickers.length === 0) {
        range.getOffsetRange(0, 1).clear();
        return allWordPickers;
      }

      return availableWordPickers;
    }


    function selectWordPickerOfTheDay(): wordPicker {
        const wordPickerOfTheDay = getAvailableWordPickers()[Math.floor(Math.random() * getAvailableWordPickers().length)];
        updateDutyValue(wordPickerOfTheDay);

        return wordPickerOfTheDay;
    }

    function updateDutyValue(wordPickerOfTheDay: wordPicker): void {
        const foundCell = range.find(wordPickerOfTheDay.email, null);
        foundCell.getOffsetRange(0, 1).setValue(true);
    }

    const selectedWordPickerOfTheDayEmail: string = selectWordPickerOfTheDay().email;

    return selectedWordPickerOfTheDayEmail;
}
