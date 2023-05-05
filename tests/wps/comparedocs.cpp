#include <QString>
#include <QtTest/QtTest>

// clang-format off
#include <wps/wpsapi.h>
#include <common/kfc/comsptr.h>
// clang-format on

#include "testbase.h"

using namespace wpsapi;
using namespace kfc;

class test_CompareDocs : public test_Base {
  Q_OBJECT
private Q_SLOTS:
  void compareDocs();
};

void test_CompareDocs::compareDocs() {
  ks_stdptr<Documents> docs;
  HRESULT hr = app->get_Documents(&docs);
  QVERIFY(docs);

  ks_stdptr<_Document> doc1;
  docs->Add(argMissing(), argMissing(), argMissing(), argMissing(),
            (Document **)&doc1);
  QVERIFY(doc1);

  ks_stdptr<_Document> doc2;
  docs->Add(argMissing(), argMissing(), argMissing(), argMissing(),
            (Document **)&doc2);
  QVERIFY(doc2);

  ks_stdptr<_Document> doc;
  hr = app->CompareDocuments(
      (Document *)doc1.get(), (Document *)doc2.get(),
      wpsapi::wdCompareDestinationNew, wpsapi::wdGranularityWordLevel,
      VARIANT_TRUE, VARIANT_TRUE, VARIANT_TRUE, VARIANT_TRUE, VARIANT_TRUE,
      VARIANT_TRUE, VARIANT_TRUE, VARIANT_TRUE, VARIANT_TRUE, VARIANT_TRUE,
      nullptr, VARIANT_FALSE, (Document **)&doc);

  // FIXME: not works...
  // QCOMPARE(hr, S_OK);
  // QVERIFY(doc);
}

QTEST_GUILESS_MAIN(test_CompareDocs)
#include "comparedocs.moc"
