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

using namespace etapi;
using namespace kfc;

class test_RangeValue : public test_EtBase {
  Q_OBJECT
private Q_SLOTS:
  void rangeValue();

private:
  void getRange(_Worksheet *sheet, const WCHAR *cell, IRange **range) {
    ks_bstr bstr(cell);

    VARIANT var;
    V_VT(&var) = VT_BSTR;
    V_BSTR(&var) = bstr;
    HRESULT hr = sheet->get_Range(var, *argMissing(), (Range **)range);
    QCOMPARE(hr, S_OK);
  }

  template <size_t N>
  void makeValue(int (&arr)[N], VARIANT *var) {
    SAFEARRAYBOUND sab;
    sab.lLbound = 0;
    sab.cElements = N;
    SAFEARRAY *sa = SafeArrayCreate(VT_VARIANT, 1, &sab);
    QVERIFY(sa);

    for (int i = 0; i < N; ++i) {
      VARIANT value;
      V_VT(&value) = VT_I4;
      V_I4(&value) = arr[i];
      SafeArrayPutElement(sa, &i, &value);
    }

    V_VT(var) = VT_ARRAY | VT_VARIANT;
    V_ARRAY(var) = sa;
  }

  template <size_t row, size_t col>
  void makeValue(int (&arr)[row][col], VARIANT *var) {
    SAFEARRAYBOUND sab[2];
    sab[0].lLbound = 0;
    sab[0].cElements = row;
    sab[1].lLbound = 0;
    sab[1].cElements = col;
    SAFEARRAY *sa = SafeArrayCreate(VT_VARIANT, 2, sab);
    QVERIFY(sa);

    for (int i = 0; i < row; ++i) {
      for (int j = 0; j < col; ++j) {
        VARIANT value;
        V_VT(&value) = VT_I4;
        V_I4(&value) = arr[i][j];
        INT32 indices[] = {i, j};
        SafeArrayPutElement(sa, indices, &value);
      }
    }

    V_VT(var) = VT_ARRAY | VT_VARIANT;
    V_ARRAY(var) = sa;
  }

  void checkCellValue(_Worksheet *sheet, const WCHAR *cell, int expected) {
    ks_stdptr<IRange> range;
    getRange(sheet, cell, &range);
    QVERIFY(range);

    VARIANT value;
    HRESULT hr = range->get_Value(*argMissing(), 0, &value);
    QCOMPARE(hr, S_OK);
    QCOMPARE(V_VT(&value), VT_R8);
    QCOMPARE(V_R8(&value), expected);
  }
};

void test_RangeValue::rangeValue() {
  ks_stdptr<Workbooks> workbooks;
  app->get_Workbooks(&workbooks);
  QVERIFY(workbooks);

  ks_stdptr<_Workbook> workbook;
  workbooks->Add(*argMissing(), 0, (Workbook **)&workbook);
  QVERIFY(workbook);

  ks_stdptr<_Worksheet> sheet;
  workbook->get_ActiveSheet((IDispatch **)&sheet);
  QVERIFY(sheet);

  {
    ks_stdptr<IRange> range;
    getRange(sheet, __X("A1:B1"), &range);
    QVERIFY(range);

    VARIANT value;
    int arr[] = {1, 2};
    makeValue(arr, &value);
    HRESULT hr = range->put_Value(*argMissing(), 0, value);
    QCOMPARE(hr, S_OK);

    checkCellValue(sheet, __X("A1"), 1);
    checkCellValue(sheet, __X("B1"), 2);
  }

  {
    ks_stdptr<IRange> range;
    getRange(sheet, __X("A2:C3"), &range);
    QVERIFY(range);

    VARIANT value;
    int arr[][3] = {{3, 4, 5}, {6, 7, 8}};
    makeValue(arr, &value);
    HRESULT hr = range->put_Value(*argMissing(), 0, value);
    QCOMPARE(hr, S_OK);

    checkCellValue(sheet, __X("A2"), 3);
    checkCellValue(sheet, __X("C3"), 8);
  }

  VARIANT saveChanges;
  V_VT(&saveChanges) = VT_BOOL;
  V_BOOL(&saveChanges) = VARIANT_FALSE;
  workbook->Close(saveChanges, *argMissing(), *argMissing(), 0);
}

QTEST_GUILESS_MAIN(test_RangeValue)
#include "rangevalue.moc"
