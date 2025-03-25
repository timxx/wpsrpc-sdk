#include "testbase.h"

#include <QtTest/QtTest>

// clang-format off
#include <wps/wpsapi.h>
#include <common/kfc/comsptr.h>
// clang-format on

using namespace wpsapi;
using namespace kfc;

VARIANT *argMissing() {
  static VARIANT s_varMissing;
  // reset every time, shall we make a assert?
  V_VT(&s_varMissing) = VT_ERROR;
  V_ERROR(&s_varMissing) = 0x80020004;

  return &s_varMissing;
}

void test_Base::initTestCase() {
  IKRpcClient *rpc = nullptr;
  HRESULT hr = createWpsRpcInstance(&rpc);
  QCOMPARE(hr, S_OK);
  QVERIFY(rpc != nullptr);

  hr = rpc->getWpsApplication((IUnknown **)&app);
  QVERIFY(app != nullptr);

  hr = rpc->getProcessPid(&pid);
  QCOMPARE(hr, S_OK);
  QVERIFY(pid > 0);
}

void test_Base::cleanupTestCase() {
  app->Quit(argMissing(), argMissing(), argMissing());
  app->Release();

  // we can't free the rpc, otherwise it crashes...
  // delete rpc;

  QThread::msleep(100);

  // kill the process, as quit may not really work
  system(QString("pkill -P %1").arg(pid).toUtf8().constData());
}
