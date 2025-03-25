#include "testetbase.h"

#include <QtTest/QtTest>
#include <QCoreApplication>

#include <et/etapi.h>

using namespace etapi;
using namespace kfc;

VARIANT *argMissing() {
  static VARIANT s_varMissing;
  // reset every time, shall we make a assert?
  V_VT(&s_varMissing) = VT_ERROR;
  V_ERROR(&s_varMissing) = 0x80020004;

  return &s_varMissing;
}

void test_EtBase::initTestCase() {
  IKRpcClient *rpc = nullptr;
  HRESULT hr = createEtRpcInstance(&rpc);
  QCOMPARE(hr, S_OK);
  QVERIFY(rpc != nullptr);

  hr = rpc->getEtApplication((IUnknown **)&app);
  QVERIFY(app != nullptr);

  hr = rpc->getProcessPid(&pid);
  QCOMPARE(hr, S_OK);
  QVERIFY(pid > 0);
}

void test_EtBase::cleanupTestCase() {
  app->Quit();
  app->Release();

  QThread::msleep(100);

  // kill the process, as quit may not really work
  system(QString("pkill -P %1").arg(pid).toUtf8().constData());
}

ks_bstr test_EtBase::getDataFile(const QString &filename) {
  QString fullPath = qApp->applicationDirPath() + "/data/" + filename;
  return ks_bstr(fullPath.utf16());
}

void test_EtBase::getRange(_Worksheet *sheet, const WCHAR *cell, IRange **range) {
  ks_bstr bstr(cell);

  VARIANT var;
  V_VT(&var) = VT_BSTR;
  V_BSTR(&var) = bstr;
  HRESULT hr = sheet->get_Range(var, *argMissing(), (Range **)range);
  QCOMPARE(hr, S_OK);
}