import 'package:excel/excel.dart';
import 'package:flutter/material.dart';

class Shop {
  Shop(
    this.shopID, {
    required this.shopName,
    this.updateDate,
    this.status,
  });

  String shopID;
  String shopName;
  DateTime? updateDate;
  String? status;
}

final shops = [
  Shop(
    'SHOP01',
    shopName: 'Shop need name1',
    updateDate: DateTime.now(),
    status: 'เสนอราคา',
  ),
  Shop(
    'SHOP02',
    shopName: 'Shop need name2',
    updateDate: DateTime.now(),
    status: 'ขอราคา',
  ),
  null,
  Shop(
    'SHOP04',
    shopName: 'Shop need name4',
    updateDate: DateTime.now(),
    status: 'ขอราคา',
  ),
  Shop(
    'SHOP05',
    shopName: 'Shop need name5',
    updateDate: DateTime.now(),
    status: 'เสนอราคา',
  ),
];

const itemSection = [
  'ลำดับ',
  'รหัส',
  'ชื่อ',
  'แจ้งลบ',
];

const partsTypeSection = [
  'แท้',
  'เทียม',
  'เก่า',
  'สั่ง',
];

class ItemInfo {
  ItemInfo({
    required this.itemID,
    this.order,
    this.itemCode,
    required this.itemName,
    this.deleteNotice,
  });

  String itemID;
  String? order;
  String? itemCode;
  String itemName;
  String? deleteNotice;
}

final itemInfoList = [
  ItemInfo(order: '1', itemID: 'ITEM1', itemCode: '001', itemName: 'ITEM1', deleteNotice: 'แจ้ง'),
  ItemInfo(order: '2', itemID: 'ITEM2', itemCode: '002', itemName: 'ITEM2', deleteNotice: 'แจ้ง'),
  ItemInfo(itemID: 'ITEM3', itemName: 'ITEM3'),
  ItemInfo(itemID: 'ITEM4', itemName: 'ITEM5'),
  ItemInfo(itemID: 'ITEM6', itemName: 'ITEM6'),
];

final shopPartsPrice = {
  "SHOP01": {
    'ITEM1': {
      "M": 2000.0,
      "T": null,
      "K": 500.0,
      "R": 0.0,
    },
    'ITEM2': {
      "M": 100.0,
      "T": null,
      "K": null,
      "R": 0.0,
    },
    'ITEM3': {
      "M": 1500.0,
      "T": null,
      "K": null,
      "R": 0.0,
    },
  },
  "SHOP02": {
    "ITEM1": {
      "M": 2500.0,
      "T": null,
      "K": 5000.0,
      "R": 0.0,
    },
    "ITEM2": {
      "M": 120.0,
      "T": null,
      "K": null,
      "R": 0.0,
    },
    "ITEM3": {
      "M": 1200.0,
      "T": null,
      "K": null,
      "R": null,
    },
  },
  "SHOP04": {
    "ITEM1": {
      "M": 2500.0,
      "T": null,
      "K": 5000.0,
      "R": 0.0,
    },
    "ITEM2": {
      "M": 1200.0,
      "T": null,
      "K": null,
      "R": null,
    },
  },
};

final partsType = ['M', 'T', 'K', 'R'];

void exportExcel() {
  var excel = Excel.createExcel();

  var sheet = excel['Example'];
  excel.delete('Sheet1');

  var cellStyle = CellStyle(fontSize: 10, horizontalAlign: HorizontalAlign.Center);

  //------------------------ Shop
  // start at row 0 col 4 (E1)
  int rowShopIndex = 0;
  int colShopIndex = 4;

  // 5 time
  for (var shop in shops) {
    var list = [shop?.shopName, shop?.updateDate, shop?.status];
    // 3 time
    for (var value in list) {
      var cellShopIndexStart = CellIndex.indexByColumnRow(
        rowIndex: rowShopIndex,
        columnIndex: colShopIndex,
      );

      var cellShopIndexEnd = CellIndex.indexByColumnRow(
        rowIndex: rowShopIndex,
        columnIndex: colShopIndex + 3,
      );

      sheet.merge(cellShopIndexStart, cellShopIndexEnd);
      sheet.cell(cellShopIndexStart)
        ..value = value ?? '-'
        ..cellStyle = cellStyle;

      rowShopIndex++;
    }
    rowShopIndex = 0;
    colShopIndex += 4;
  }

  //------------------------ Header
  // start at row 3 col 0 (A4)
  int rowHeaderIndex = 3;
  int colHeaderIndex = 0;

  for (var item in itemSection) {
    // 4 time
    sheet.cell(CellIndex.indexByColumnRow(
      rowIndex: rowHeaderIndex,
      columnIndex: colHeaderIndex,
    ))
      ..value = item
      ..cellStyle = cellStyle;
    colHeaderIndex++;
  }

  // 5 time
  List.generate(5, (index) {
    // 4 time
    for (var type in partsTypeSection) {
      sheet.cell(CellIndex.indexByColumnRow(
        rowIndex: rowHeaderIndex,
        columnIndex: colHeaderIndex,
      ))
        ..value = type
        ..cellStyle = cellStyle;
      colHeaderIndex++;
    }
  });

  //------------------------ Data

  // 1-∞ time
  for (var item in itemInfoList) {
    List<dynamic> list = [item.order, item.itemCode, item.itemName, item.deleteNotice];
    var itemID = item.itemID;
    // 5 time
    for (var shop in shops) {
      var shopID = shop?.shopID;

      // 4 time
      for (var type in partsType) {
        var cost = shopPartsPrice[shopID]?[itemID]?[type];
        list.add(cost);
      }
    }
    sheet.appendRow(list);
  }

  //------------------------ Summary
  int rowSumIndex = sheet.maxRows;

  sheet.merge(CellIndex.indexByColumnRow(rowIndex: rowSumIndex, columnIndex: 0),
      CellIndex.indexByColumnRow(rowIndex: rowSumIndex, columnIndex: 3));
  sheet.cell(CellIndex.indexByColumnRow(rowIndex: rowSumIndex, columnIndex: 0))
    ..value = 'ราคารวมอะไหล่แต่ละประเภท'
    ..cellStyle = cellStyle;

  Map<String?, Map<String, double?>> sumResult = {};
  Map<String, double?> emptyResult = {'M': 0.0, 'T': 0.0, 'K': 0.0, 'R': 0.0};

  for (var shop in shops) {
    if (shop == null) {
      sumResult.addAll({null: emptyResult});
      continue;
    }

    final shopID = shop.shopID;

    if (!shopPartsPrice.containsKey(shopID)) {
      sumResult[shopID] = emptyResult;
      continue;
    }

    sumResult[shopID] = {};

    final items = shopPartsPrice[shopID];
    final itemKeys = items!.keys.toList();

    for (final key in partsType) {
      final values = itemKeys.map((item) => items[item]?[key] ?? 0).toList();

      sumResult[shopID]![key] = values.reduce((a, b) => a + b);
    }
  }
  print(sumResult);

  //plot excel
  int colSumIndex = 4;

  sumResult.forEach((shop, costs) {
    for (var cost in costs.values) {
      sheet.cell(CellIndex.indexByColumnRow(rowIndex: rowSumIndex, columnIndex: colSumIndex++))
        ..value = cost
        ..cellStyle = CellStyle(fontSize: 10);
    }
  });

  excel.save(fileName: 'test.xlsx');
}

class ExcelExample extends StatelessWidget {
  const ExcelExample({super.key});

  @override
  Widget build(BuildContext context) {
    return const Scaffold(
      body: Center(
        child: Text('ExcelExample'),
      ),
      floatingActionButton: FloatingActionButton(
        onPressed: exportExcel,
        tooltip: 'Export',
        child: Icon(Icons.downloading),
      ), //
    );
  }
}
