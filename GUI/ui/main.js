let settings = {
    'excelPath': '',
    'driverPath': '',
    'locPath': ''
}

async function LoadLocFileAndUpdateDropdown() {
    let locationsText = await eel.askForLocFile()()

    let locationsData = locationsText.split('\n')

    let locationDD = document.querySelector('#locationDropdown')

    for (let i = 0; i < locationsData.length; i++) {
        let option = document.createElement('option')
        option.innerHTML = locationsData[i].trim()
        locationDD.appendChild(option)
    }
}

async function SelectExcelPath() {
    let excelPath = await eel.askForExcelFile()()

    if (excelPath === 'error') { alert('Выбран файл не с расширением Excel (расширения .xlsx, .slsm, .xltx, .xltm)') }


}
