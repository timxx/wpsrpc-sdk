#ifndef WPS_RPC_SDK_TEST_BASE_H
#define WPS_RPC_SDK_TEST_BASE_H

#include <QObject>
#include <qobjectdefs.h>

namespace wpsapi {
struct _Application;
}

typedef struct tagVARIANT VARIANT;
VARIANT *argMissing();

class test_Base : public QObject {
  Q_OBJECT
protected Q_SLOTS:
  void initTestCase();
  void cleanupTestCase();

protected:
  wpsapi::_Application *app = nullptr;
};

#endif // WPS_RPC_SDK_TEST_BASE_H
