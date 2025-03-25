#ifndef WPS_RPC_SDK_TEST_ETBASE_H
#define WPS_RPC_SDK_TEST_ETBASE_H

#include <QObject>
#include "comdef.h"

namespace etapi {
struct _Application;
}

typedef struct tagVARIANT VARIANT;
VARIANT *argMissing();

class test_EtBase : public QObject {
  Q_OBJECT
protected Q_SLOTS:
  void initTestCase();
  void cleanupTestCase();

protected:
  kfc::ks_bstr getDataFile(const QString &filename);
  void getRange(etapi::_Worksheet *sheet, const WCHAR *cell, etapi::IRange **range);

protected:
  etapi::_Application *app = nullptr;
  long long pid = 0;
};

#endif // WPS_RPC_SDK_TEST_ETBASE_H
