#include "testetbase.h"

#include <QtTest/QtTest>

#include <et/etapi.h>

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
}

void test_EtBase::cleanupTestCase() {
  app->Quit();
  app->Release();
}
