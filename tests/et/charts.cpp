#include <QString>
#include <QtTest/QtTest>

// clang-format off
#include <et/etapi.h>
#include <common/kfc/comsptr.h>
#include <qtestcase.h>
// clang-format on

#include "comdef.h"
#include "oaidl.h"
#include "oleauto.h"
#include "testetbase.h"
#include "variant.h"

using namespace etapi;
using namespace kfc;

class test_Charts : public test_EtBase {
  Q_OBJECT
private Q_SLOTS:
  void testExistingCharts();
  void testExistingCharts2();
  void testAddChart();
};

void test_Charts::testExistingCharts() {
  ks_stdptr<Workbooks> workbooks;
  app->get_Workbooks(&workbooks);
  QVERIFY(workbooks);

  ks_stdptr<_Workbook> workbook;
  ks_bstr fileName = getDataFile("chart.xlsx");
  workbooks->Open(fileName, *argMissing(), *argMissing(), *argMissing(),
                  *argMissing(), *argMissing(), *argMissing(), *argMissing(),
                  *argMissing(), *argMissing(), *argMissing(), *argMissing(),
                  *argMissing(), *argMissing(), *argMissing(), 0,
                  (Workbook **)&workbook);
  QVERIFY(workbook);

  ks_stdptr<_Worksheet> sheet;
  workbook->get_ActiveSheet((IDispatch **)&sheet);
  QVERIFY(sheet);

  KComVariant index(1);
  ks_stdptr<IChartObjects> chartObjects;
  sheet->ChartObjects(index, 2052, (IDispatch **)&chartObjects);
  QVERIFY(chartObjects);

  long count = 0;
  chartObjects->get_Count(&count);
  QCOMPARE(count, 0);

  HRESULT hr = chartObjects->Delete(argMissing());
  QCOMPARE(hr, S_OK);

  VARIANT saveChanges;
  V_VT(&saveChanges) = VT_BOOL;
  V_BOOL(&saveChanges) = VARIANT_FALSE;
  workbook->Close(saveChanges, *argMissing(), *argMissing(), 0);
}

void test_Charts::testExistingCharts2() {
  ks_stdptr<Workbooks> workbooks;
  app->get_Workbooks(&workbooks);
  QVERIFY(workbooks);

  ks_stdptr<_Workbook> workbook;
  ks_bstr fileName = getDataFile("chart.xlsx");
  workbooks->Open(fileName, *argMissing(), *argMissing(), *argMissing(),
                  *argMissing(), *argMissing(), *argMissing(), *argMissing(),
                  *argMissing(), *argMissing(), *argMissing(), *argMissing(),
                  *argMissing(), *argMissing(), *argMissing(), 0,
                  (Workbook **)&workbook);
  QVERIFY(workbook);

  ks_stdptr<_Worksheet> sheet;
  workbook->get_ActiveSheet((IDispatch **)&sheet);
  QVERIFY(sheet);

  ks_stdptr<IChartObjects> chartObjects;
  sheet->ChartObjects(*argMissing(), 0, (IDispatch **)&chartObjects);
  QVERIFY(chartObjects);

  long count = 0;
  chartObjects->get_Count(&count);
  QCOMPARE(count, 1);

  VARIANT saveChanges;
  V_VT(&saveChanges) = VT_BOOL;
  V_BOOL(&saveChanges) = VARIANT_FALSE;
  workbook->Close(saveChanges, *argMissing(), *argMissing(), 0);
}

// see https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects
void test_Charts::testAddChart() {
  ks_stdptr<Workbooks> workbooks;
  app->get_Workbooks(&workbooks);
  QVERIFY(workbooks);

  ks_stdptr<_Workbook> workbook;
  workbooks->Add(*argMissing(), 0, (Workbook **)&workbook);
  QVERIFY(workbook);

  ks_stdptr<_Worksheet> sheet;
  workbook->get_ActiveSheet((IDispatch **)&sheet);
  QVERIFY(sheet);

  ks_stdptr<IChartObjects> chartObjects;
  sheet->ChartObjects(*argMissing(), 0, (IDispatch **)&chartObjects);
  QVERIFY(chartObjects);

  long count = 0;
  chartObjects->get_Count(&count);
  QCOMPARE(count, 0);

  ks_stdptr<IChartObject> chartObject;
  chartObjects->Add(100, 30, 400, 250, (ChartObject **)&chartObject);
  QVERIFY(chartObject);

  ks_stdptr<_Chart> chart;
  chartObject->get_Chart((Chart**)&chart);
  QVERIFY(chart);

  ks_stdptr<IRange> range;
  getRange(sheet, __X("A1:A20"), (IRange **)&range);
  KComVariant source(range);
  KComVariant gallery(xlLine);
  KComVariant title(__X("New Chart"));
  chart->ChartWizard(source, gallery, *argMissing(), *argMissing(),
                     *argMissing(), *argMissing(), *argMissing(), title,
                     *argMissing(), *argMissing(), *argMissing(), 0);

  ks_stdptr<IChartArea> chartArea;
  chart->get_ChartArea(0, (ChartArea**)&chartArea);
  QVERIFY(chartArea);

  ks_stdptr<IChartFormat> format;
  chartArea->get_Format((ChartFormat**)&format);
  QVERIFY(format);

  ks_stdptr<FillFormat> fillFormat;
  format->get_Fill((FillFormat**)&fillFormat);
  QVERIFY(fillFormat);

  HRESULT hr = fillFormat->Patterned(msoPatternLightDownwardDiagonal);
  QCOMPARE(hr, S_OK);

  // this is bug for wps spreadsheets
  chartObjects->get_Count(&count);
  QCOMPARE(count, 0);

  // we have to re-get the chartObjects
  chartObjects.clear();
  sheet->ChartObjects(*argMissing(), 0, (IDispatch **)&chartObjects);
  QVERIFY(chartObjects);

  chartObjects->get_Count(&count);
  QCOMPARE(count, 1);

  VARIANT saveChanges;
  V_VT(&saveChanges) = VT_BOOL;
  V_BOOL(&saveChanges) = VARIANT_FALSE;
  workbook->Close(saveChanges, *argMissing(), *argMissing(), 0);
}

QTEST_GUILESS_MAIN(test_Charts)
#include "charts.moc"
