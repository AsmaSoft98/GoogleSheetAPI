using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetApi;
using Microsoft.AspNetCore.Mvc;
using System.Linq;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace GoogleSheetAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductsController : ControllerBase
    {
        const string SPREADSHEET_ID = "1mfXO_J-ZUbmqDp_Ym45Sjbbs7mtMK_SVq21WKcbEahk";
        const string SHEET_NAME = "Items";

        SpreadsheetsResource.ValuesResource _googleSheetValues;

        public ProductsController(GoogleSheetHelper googleSheetsHelper)
        {
            _googleSheetValues = googleSheetsHelper.Service.Spreadsheets.Values;
        }

        [HttpGet]
        public IActionResult Get()
        {
            var range = $"{SHEET_NAME}!A:D";

            var request = _googleSheetValues.Get(SPREADSHEET_ID, range);
            var response = request.Execute();
            var values = response.Values;

            return Ok(ProductsMapper.MapFromRangeData(values));
        }

        [HttpGet("{rowId}")]
        public IActionResult Get(int rowId)
        {
            var range = $"{SHEET_NAME}!A{rowId}:D{rowId}";
            var request = _googleSheetValues.Get(SPREADSHEET_ID, range);
            var response = request.Execute();
            var values = response.Values;

            return Ok(ProductsMapper.MapFromRangeData(values).FirstOrDefault());
        }

        [HttpPost]
        public IActionResult Post(Product item)
        {
            var range = $"{SHEET_NAME}!A:D";
            var valueRange = new ValueRange
            {
                Values = ProductsMapper.MapToRangeData(item)
            };

            var appendRequest = _googleSheetValues.Append(valueRange, SPREADSHEET_ID, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();

            return CreatedAtAction(nameof(Get), item);
        }

        [HttpPut("{rowId}")]
        public IActionResult Put(int rowId, Product item)
        {
            var range = $"{SHEET_NAME}!A{rowId}:D{rowId}";
            var valueRange = new ValueRange
            {
                Values = ProductsMapper.MapToRangeData(item)
            };

            var updateRequest = _googleSheetValues.Update(valueRange, SPREADSHEET_ID, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();

            return NoContent();
        }

        [HttpDelete("{rowId}")]
        public IActionResult Delete(int rowId)
        {
            var range = $"{SHEET_NAME}!A{rowId}:D{rowId}";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = _googleSheetValues.Clear(requestBody, SPREADSHEET_ID, range);
            deleteRequest.Execute();

            return NoContent();
        }
    }
}
