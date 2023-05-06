#ifndef WPS_RPC_SDK_TEST_ETBASE_H
#define WPS_RPC_SDK_TEST_ETBASE_H

#include <QObject>

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
  etapi::_Application *app = nullptr;
};

#endif // WPS_RPC_SDK_TEST_ETBASE_H
