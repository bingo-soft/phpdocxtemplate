<?php

namespace Doctrine\Tests\DBAL\Query;

use PHPUnit\Framework\TestCase;
use DOMDocument;
use PhpDocxTemplate\PhpDocxTemplate;
use PhpDocxTemplate\DocxDocument;
use Twig\Loader\ArrayLoader;
use Twig\Environment;

class PhpDocxTemplateTest extends TestCase
{
    private const TEMPLATE1 = __DIR__ . "/templates/template1.docx";
    private const TEMPLATE2 = __DIR__ . "/templates/template2.docx";
    private const TEMPLATE3 = __DIR__ . "/templates/template3.docx";
    private const TEMPLATE4 = __DIR__ . "/templates/template4.docx";

    public function testXmlToString(): void
    {
        $xml = new DOMDocument('1.0');
        $root = $xml->createElement('book');
        $root = $xml->appendChild($root);
        $title = $xml->createElement('title');
        $title = $root->appendChild($title);
        $text = $xml->createTextNode('Title');
        $title->appendChild($text);
        $reporter = new PhpDocxTemplate(self::TEMPLATE1);

        $this->assertEquals(
            $reporter->xmlToString($xml),
            "<?xml version=\"1.0\"?>\n<book><title>Title</title></book>\n"
        );
        $reporter->close();
    }

    public function testGetDocx(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE1);
        $this->assertInstanceOf(DocxDocument::class, $reporter->getDocx());
        $reporter->close();
    }

    public function testGetXml(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE1);
        $this->assertEquals(
            $reporter->getXml(),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" .
            "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" " .
            "xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" " .
            "xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" " .
            "xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" " .
            "xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" " .
            "xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" " .
            "xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" " .
            "xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" " .
            "xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" " .
            "xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" " .
            "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " .
            "xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" " .
            "xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" " .
            "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " .
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " .
            "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" " .
            "xmlns:v=\"urn:schemas-microsoft-com:vml\" " .
            "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" " .
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " .
            "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " .
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " .
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " .
            "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" " .
            "xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" " .
            "xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" " .
            "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" " .
            "xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" " .
            "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" " .
            "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " .
            "mc:Ignorable=\"w14 w15 w16se w16cid wp14\"><w:body><w:p w14:paraId=\"504F2588\" " .
            "w14:textId=\"54DF26C8\" w:rsidR=\"0090657C\" w:rsidRPr=\"00FA3F61\" " .
            "w:rsidRDefault=\"00FA3F61\" w:rsidP=\"00C13DD6\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/>" .
            "</w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>Hello {{ object }}!</w:t>" .
            "</w:r><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"0\"/></w:p>" .
            "<w:sectPr w:rsidR=\"0090657C\" w:rsidRPr=\"00FA3F61\"><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" .
            "<w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" " .
            "w:footer=\"708\" w:gutter=\"0\"/><w:cols w:space=\"708\"/><w:docGrid w:linePitch=\"360\"/>" .
            "</w:sectPr></w:body></w:document>\n"
        );
        $reporter->close();
    }

    public function testPatchXml(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE1);
        //test stripTags
        $xml = "{<tag>%Hello</w:t><w:t>\nworld%<tag>}\n{<tag>{Hi</w:t><w:t>\nthere}<tag>}\n";
        $this->assertEquals(
            $reporter->patchXml($xml),
            "{%Hello\nworld%}\n{{Hi\nthere}}\n"
        );

        //test colspan
        $xml = "<w:tc xeLLm[t6;cT&!Z_#KI8cniins[)UX>TAnAaqg_a}sePvK.OO#Q=B-]cBDFM8UL]8m@i" .
               "Ct{% colspan val%}TkuSd<w:r meg+PYSJWO}~k<w:t></w:t></w:r>" .
               "<w:gridSpan88MJ@1bX/><w:tcPrL4><w:gridSpan@1bY/>?Nl`z:^kY@FXeJ@P{8WhCt0__/,8woI2." .
               "8#[r_Cqig!5Qt{8gl5ls<9Ci|^QN2IK#L[cB9@:XclVQQIxe</w:tc>";
        $this->assertEquals(
            $reporter->patchXml($xml),
            '<w:tc xeLLm[t6;cT&!Z_#KI8cniins[)UX>TAnAaqg_a}sePvK.OO#Q=B-]cBDFM8UL]8m@iCtTkuSd<w:tcPrL4>' .
            '<w:gridSpan w:val="{{val}}"/><w:gridSpan@1bY/>?Nl`z:^kY@FXeJ@P{8WhCt0__/,8woI2.8#[r_Cqig!5Qt' .
            '{8gl5ls<9Ci|^QN2IK#L[cB9@:XclVQQIxe</w:tc>'
        );

        //test cellbg
        $xml = "<w:tc xeLLm[t6;cT&!Z_#KI8cniins[)UX>TAnAaqg_a}sePvK.OO#Q=B-]cBDFM8UL]8m@i" .
               "Ct{% cellbg val%}TkuSd<w:r meg+PYSJWO}~k<w:t></w:t></w:r>" .
               "<w:shd88MJ@1bX/><w:tcPrL4><w:shd@1bY/>?Nl`z:^kY@FXeJ@P{8WhCt0__/,8woI2." .
               "8#[r_Cqig!5Qt{8gl5ls<9Ci|^QN2IK#L[cB9@:XclVQQIxe</w:tc>";
        $this->assertEquals(
            $reporter->patchXml($xml),
            '<w:tc xeLLm[t6;cT&!Z_#KI8cniins[)UX>TAnAaqg_a}sePvK.OO#Q=B-]cBDFM8UL]8m@iCtTkuSd<w:tcPrL4>' .
            '<w:shd w:val="clear" w:color="auto" w:fill="{{val}}"/><w:shd@1bY/>?Nl`z:^kY@FXeJ@P{8WhCt0__/,' .
            '8woI2.8#[r_Cqig!5Qt{8gl5ls<9Ci|^QN2IK#L[cB9@:XclVQQIxe</w:tc>'
        );

        //test avoid {{r and {%r tags
        $xml = "<w:t>[x&\P^XIk]sdD((KK{%r  .S-ce4F!b,l8Qo`(>NA;%}";
        $this->assertEquals(
            $reporter->patchXml($xml),
            '<w:t xml:space="preserve">[x&\P^XIk]sdD((KK</w:t></w:r><w:r><w:t xml:space="preserve">' .
            '{%r  .S-ce4F!b,l8Qo`(>NA;%}</w:t></w:r><w:r><w:t xml:space="preserve">'
        );
        $xml = "{%r _Rom{X=aC3/s#W#~o<#d:tH^>DTAz;s<}O0RJ#V!wW:]%kR@wzLf*\iu^zAGrr!3]v<SUc|B)o>kA.:*1?,0%}";
        $this->assertEquals(
            $reporter->patchXml($xml),
            '</w:t></w:r><w:r><w:t xml:space="preserve">{%r _Rom{X=aC3/s#W#~o<#d:tH^>DTAz;s<}O0RJ#V!wW:]%kR' .
            '@wzLf*\iu^zAGrr!3]v<SUc|B)o>kA.:*1?,0%}</w:t></w:r><w:r><w:t xml:space="preserve">'
        );

        // test
        $xml = "<w:trs>RaYI}@{{trs :k!HJO Z36#T1$3U>6F.=y1_y5w:I7uAWs=_n-(ix*HXPd7){>{V|yPkDIa=" .
               "cLlQwozlcjFn<_(`&32 PZ1e5_Pqbc@zFR!r0(%}S'orf,78A<S nR=E</w:trs>";
        $this->assertEquals(
            $reporter->patchXml($xml),
            '{{ :k!HJO Z36#T1$3U>6F.=y1_y5w:I7uAWs=_n-(ix*HXPd7){>{V|yPkDIa=cLlQwozlcjFn<_(`&32 PZ1e5_Pqbc@zFR!r0(%}'
        );

        // test vMerge
        $xml = "<w:tc></w:tcPr>t/H-Q.X)jC_sI6(J7w-;QI&JpDG}:>f02Zls<8(7&SEyc>" .
               "`@P/<Ero^KEbL`EX^<w:t>{% vm %}</w:t></w:tc>";
        $this->assertEquals(
            $reporter->patchXml($xml),
            '<w:tc><w:vMerge w:val="{% if loop.first %}restart{% else %}continue' .
            '{% endif %}"/></w:tcPr>t/H-Q.X)jC_sI6(J7w-;QI&JpDG}:>f02Zls<8(7&SEyc>`' .
            '@P/<Ero^KEbL`EX^<w:t>{% if loop.first %}{% endif %}</w:t></w:tc>'
        );

        // test hMerge
        $xml = "<w:tc></w:tcPr>t/H-Q.X)jC_sI6(J7w-;QI&JpDG}:>f02Zls<8(7&SEyc>" .
               "`@P/<Ero^KEbL`EX^<w:t>{% hm %}</w:t></w:tc>";
        $this->assertEquals(
            $reporter->patchXml($xml),
            '{% if loop.first %}<w:tc><w:gridSpan w:val="{{ loop.length }}"/></w:tcPr>t/H-Q.X)' .
            'jC_sI6(J7w-;QI&JpDG}:>f02Zls<8(7&SEyc>`@P/<Ero^KEbL`EX^<w:t></w:t></w:tc>{% endif %}'
        );

        // test cleanTags
        $xml = '{%&#8216;&lt;&gt;“”‘’%}';
        $this->assertEquals(
            $reporter->patchXml($xml),
            "{%'<>\"\"''%}"
        );
        $reporter->close();
    }

    public function testRenderXml(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE1);
        $this->assertEquals(
            $reporter->buildXml(["object" => "world"]),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" .
            "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" " .
            "xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" " .
            "xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" " .
            "xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" " .
            "xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" " .
            "xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" " .
            "xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" " .
            "xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" " .
            "xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" " .
            "xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" " .
            "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " .
            "xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" " .
            "xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" " .
            "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " .
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " .
            "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" " .
            "xmlns:v=\"urn:schemas-microsoft-com:vml\" " .
            "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" " .
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " .
            "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " .
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " .
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " .
            "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" " .
            "xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" " .
            "xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" " .
            "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" " .
            "xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" " .
            "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" " .
            "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " .
            "mc:Ignorable=\"w14 w15 w16se w16cid wp14\"><w:body><w:p w14:paraId=\"504F2588\" " .
            "w14:textId=\"54DF26C8\" w:rsidR=\"0090657C\" w:rsidRPr=\"00FA3F61\" " .
            "w:rsidRDefault=\"00FA3F61\" w:rsidP=\"00C13DD6\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/>" .
            "</w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>Hello world!</w:t>" .
            "</w:r><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"0\"/></w:p>" .
            "<w:sectPr w:rsidR=\"0090657C\" w:rsidRPr=\"00FA3F61\"><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" .
            "<w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" " .
            "w:footer=\"708\" w:gutter=\"0\"/><w:cols w:space=\"708\"/><w:docGrid w:linePitch=\"360\"/>" .
            "</w:sectPr></w:body></w:document>\n"
        );
        $reporter->close();
    }

    public function testRender(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE2);
        $reporter->render(["one" => "1", "two" => "2", "three" => "3", "four" => "4"]);
        $this->assertEquals(
            $reporter->getXml(),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" .
            "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:c" .
            "x=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft." .
            "com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/1" .
            "0/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=" .
            "\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microso" .
            "ft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/20" .
            "16/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:" .
            "cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.open" .
            "xmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/" .
            "2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:sch" .
            "emas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/r" .
            "elationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:" .
            "schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessin" .
            "gDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns" .
            ":w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordproce" .
            "ssingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"h" .
            "ttp://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/of" .
            "fice/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex" .
            "\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http:" .
            "//schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.co" .
            "m/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessing" .
            "Shape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\"><w:body><w:tbl><w:tblPr><w:tblStyle w:val=\"a3\"/" .
            "><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:fir" .
            "stColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr><w:tblGrid><w:gridCol w" .
            ":w=\"3115\"/><w:gridCol w:w=\"3115\"/><w:gridCol w:w=\"3115\"/></w:tblGrid><w:tr w:rsidR=\"00031864" .
            "\" w14:paraId=\"73B274FD\" w14:textId=\"77777777\" w:rsidTr=\"00135B64\"><w:tc><w:tcPr><w:tcW w:w=\"3" .
            "115\" w:type=\"dxa\"/></w:tcPr><w:p w14:paraId=\"29117FB3\" w14:textId=\"713E58B3\" w:rsidR=\"000318" .
            "64\" w:rsidRPr=\"0033062B\" w:rsidRDefault=\"00031864\" w:rsidP=\"0033062B\"><w:pPr><w:rPr><w:lang w" .
            ":val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>1</w:t></w:r></w:p" .
            "></w:tc><w:tc><w:tcPr><w:tcW w:w=\"6230\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/></w:tcPr><w:p w14" .
            ":paraId=\"4620CF03\" w14:textId=\"7FD6FF29\" w:rsidR=\"00031864\" w:rsidRPr=\"00B6314C\" w:rsidRDefa" .
            "ult=\"00B6314C\" w:rsidP=\"0033062B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:rPr></w:pPr><w:r><w" .
            ":rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>1</w:t></w:r></w:p></w:tc></w:tr><w:tr w:rsidR=\"00031864" .
            "\" w14:paraId=\"638CAE21\" w14:textId=\"77777777\" w:rsidTr=\"0033062B\"><w:tc><w:tcPr><w:tcW w:w=\"" .
            "3115\" w:type=\"dxa\"/><w:vMerge w:val=\"restart\"/></w:tcPr><w:p w14:paraId=\"69D3958C\" w14:textId" .
            "=\"77777777\" w:rsidR=\"00031864\" w:rsidRPr=\"0033062B\" w:rsidRDefault=\"00031864\" w:rsidP=\"0033" .
            "062B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/><" .
            "/w:rPr><w:t>2</w:t></w:r></w:p><w:p w14:paraId=\"203EF204\" w14:textId=\"751D091A\" w:rsidR=\"000318" .
            "64\" w:rsidRPr=\"0033062B\" w:rsidRDefault=\"00031864\" w:rsidP=\"0033062B\"><w:pPr><w:rPr><w:lang w" .
            ":val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>3</w:t></w:r></w:p" .
            "></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3115\" w:type=\"dxa\"/></w:tcPr><w:p w14:paraId=\"29D750B5\" w14:" .
            "textId=\"52111D71\" w:rsidR=\"00031864\" w:rsidRPr=\"00B6314C\" w:rsidRDefault=\"00B6314C\" w:rsidP=" .
            "\"0033062B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-U" .
            "S\"/></w:rPr><w:t>2</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3115\" w:type=\"dxa\"/></w:tc" .
            "Pr><w:p w14:paraId=\"443FEDA5\" w14:textId=\"77777777\" w:rsidR=\"00031864\" w:rsidRDefault=\"000318" .
            "64\" w:rsidP=\"0033062B\"/></w:tc></w:tr><w:tr w:rsidR=\"00031864\" w14:paraId=\"3400FAC6\" w14:text" .
            "Id=\"77777777\" w:rsidTr=\"0033062B\"><w:tc><w:tcPr><w:tcW w:w=\"3115\" w:type=\"dxa\"/><w:vMerge/><" .
            "/w:tcPr><w:p w14:paraId=\"27010401\" w14:textId=\"299C3CC2\" w:rsidR=\"00031864\" w:rsidRPr=\"003306" .
            "2B\" w:rsidRDefault=\"00031864\" w:rsidP=\"0033062B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:rPr" .
            "></w:pPr></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3115\" w:type=\"dxa\"/></w:tcPr><w:p w14:paraId=\"6" .
            "D034FEF\" w14:textId=\"6465FA34\" w:rsidR=\"00031864\" w:rsidRPr=\"00B6314C\" w:rsidRDefault=\"00B63" .
            "14C\" w:rsidP=\"0033062B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lan" .
            "g w:val=\"en-US\"/></w:rPr><w:t>3</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3115\" w:type=" .
            "\"dxa\"/></w:tcPr><w:p w14:paraId=\"331E7F28\" w14:textId=\"77777777\" w:rsidR=\"00031864\" w:rsidRDe" .
            "fault=\"00031864\" w:rsidP=\"0033062B\"/></w:tc></w:tr><w:tr w:rsidR=\"0033062B\" w14:paraId=\"489E0" .
            "54E\" w14:textId=\"77777777\" w:rsidTr=\"0033062B\"><w:tc><w:tcPr><w:tcW w:w=\"3115\" w:type=\"dxa\"" .
            "/></w:tcPr><w:p w14:paraId=\"24A13E38\" w14:textId=\"18E3B4BC\" w:rsidR=\"0033062B\" w:rsidRPr=\"003" .
            "3062B\" w:rsidRDefault=\"0033062B\" w:rsidP=\"0033062B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:" .
            "rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>4</w:t></w:r></w:p></w:tc><w:tc><w:tcP" .
            "r><w:tcW w:w=\"3115\" w:type=\"dxa\"/></w:tcPr><w:p w14:paraId=\"58DCBE56\" w14:textId=\"5D680B0D\" " .
            "w:rsidR=\"0033062B\" w:rsidRPr=\"00B6314C\" w:rsidRDefault=\"00B6314C\" w:rsidP=\"0033062B\"><w:pPr>" .
            "<w:rPr><w:lang w:val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>4<" .
            "/w:t></w:r><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"0\"/></w:p></w:tc><" .
            "w:tc><w:tcPr><w:tcW w:w=\"3115\" w:type=\"dxa\"/></w:tcPr><w:p w14:paraId=\"277EDAD6\" w14:textId=\"" .
            "77777777\" w:rsidR=\"0033062B\" w:rsidRDefault=\"0033062B\" w:rsidP=\"0033062B\"/></w:tc></w:tr></w:" .
            "tbl><w:p w14:paraId=\"504F2588\" w14:textId=\"658244D8\" w:rsidR=\"0090657C\" w:rsidRPr=\"0033062B\"" .
            " w:rsidRDefault=\"0090657C\" w:rsidP=\"0033062B\"/><w:sectPr w:rsidR=\"0090657C\" w:rsidRPr=\"003306" .
            "2B\"><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" " .
            "w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/><w:cols w:space=\"708\"/><w:docGri" .
            "d w:linePitch=\"360\"/></w:sectPr></w:body></w:document>\n"
        );
        $reporter->save("./tmp/test.docx");
    }

    public function testLineBreak(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE3);
        $reporter->render(["один" => "значение с \n переносом строки"]);
        $this->assertEquals(
            $reporter->getXml(),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" .
            "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:c" .
            "x=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft." .
            "com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/1" .
            "0/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=" .
            "\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microso" .
            "ft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/20" .
            "16/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:" .
            "cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.open" .
            "xmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/" .
            "2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:sch" .
            "emas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/r" .
            "elationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:" .
            "schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessin" .
            "gDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns" .
            ":w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordproce" .
            "ssingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"h" .
            "ttp://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/of" .
            "fice/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex" .
            "\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http:" .
            "//schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.co" .
            "m/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessing" .
            "Shape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\"><w:body><w:p w14:paraId=\"504F2588\" w14:textId=" .
            "\"0B366F38\" w:rsidR=\"0090657C\" w:rsidRPr=\"00BC38E6\" w:rsidRDefault=\"00BC38E6\" w:rsidP=\"003306" .
            "2B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w" .
            ":rPr><w:t xml:space=\"preserve\">значение с </w:t></w:r></w:p><w:p><w:r><w:t xml:space=\"preserve\">".
            " переносом строки</w:t></w:r><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/>" .
            "<w:bookmarkEnd w:id=\"0\"/></w:p><w:sectPr w:rsidR=\"0090657C\" w:rsidRPr=\"00BC38E6\"><w:pgSz w:w=\"" .
            "11906\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:he" .
            "ader=\"708\" w:footer=\"708\" w:gutter=\"0\"/><w:cols w:space=\"708\"/><w:docGrid w:linePitch=\"360\"/><" .
            "/w:sectPr></w:body></w:document>\n"
        );
    }

    public function testCyrillic(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE3);
        $reporter->render(["один" => "значение"]);
        $this->assertEquals(
            $reporter->getXml(),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" .
            "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:c" .
            "x=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft." .
            "com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/1" .
            "0/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=" .
            "\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microso" .
            "ft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/20" .
            "16/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:" .
            "cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.open" .
            "xmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/" .
            "2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:sch" .
            "emas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/r" .
            "elationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:" .
            "schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessin" .
            "gDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns" .
            ":w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordproce" .
            "ssingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"h" .
            "ttp://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/of" .
            "fice/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex" .
            "\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http:" .
            "//schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.co" .
            "m/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessing" .
            "Shape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\"><w:body><w:p w14:paraId=\"504F2588\" w14:textId=" .
            "\"0B366F38\" w:rsidR=\"0090657C\" w:rsidRPr=\"00BC38E6\" w:rsidRDefault=\"00BC38E6\" w:rsidP=\"003306" .
            "2B\"><w:pPr><w:rPr><w:lang w:val=\"en-US\"/></w:rPr></w:pPr><w:r><w:rPr><w:lang w:val=\"en-US\"/></w" .
            ":rPr><w:t xml:space=\"preserve\">значение</w:t></w:r><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/>" .
            "<w:bookmarkEnd w:id=\"0\"/></w:p><w:sectPr w:rsidR=\"0090657C\" w:rsidRPr=\"00BC38E6\"><w:pgSz w:w=\"" .
            "11906\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:he" .
            "ader=\"708\" w:footer=\"708\" w:gutter=\"0\"/><w:cols w:space=\"708\"/><w:docGrid w:linePitch=\"360\"/><" .
            "/w:sectPr></w:body></w:document>\n"
        );
    }

    public function testForLoop(): void
    {
        $reporter = new PhpDocxTemplate(self::TEMPLATE4);
        $reporter->render(["сотрудники" => [
            [
                "фамилия" => "Иванов",
                "имя" => "Иван",
                "отчество" => "Иванович",
                "дети" => [
                    [
                        "фамилия" => "Иванова",
                        "имя" => "Алена",
                        "отчество" => "Ивановна",
                        "возраст" => 25
                    ],
                    [
                        "фамилия" => "Иванов",
                        "имя" => "Михаил",
                        "отчество" => "Иванович",
                        "возраст" => 6
                    ]
                ],
                "возраст" => 50
            ],
            [
                "фамилия" => "Петров",
                "имя" => "Петр",
                "отчество" => "Петрович",
                "возраст" => 30
            ]
        ]]);
        $this->assertEquals(
            $reporter->getXml(),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" .
            "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:c" .
            "x=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft." .
            "com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/1" .
            "0/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=" .
            "\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microso" .
            "ft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/20" .
            "16/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:" .
            "cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.open" .
            "xmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/" .
            "2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:sch" .
            "emas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/r" .
            "elationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:" .
            "schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessin" .
            "gDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns" .
            ":w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordproce" .
            "ssingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"h" .
            "ttp://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cid=\"http://schemas.microsoft.com/of" .
            "fice/word/2016/wordml/cid\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex" .
            "\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http:" .
            "//schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.co" .
            "m/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessing" .
            "Shape\" mc:Ignorable=\"w14 w15 w16se w16cid wp14\"><w:body><w:p w14:paraId=\"45D5C362\" w14:textId=" .
            "\"79C82D83\" w:rsidR=\"00AF4B5B\" w:rsidRDefault=\"005D75A1\" w:rsidP=\"006B0D50\"><w:pPr><w:pStyle w" .
            ":val=\"a4\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:proofErr w:typ" .
            "e=\"gramStart\"/><w:r w:rsidRPr=\"005D75A1\"><w:t xml:space=\"preserve\">Иванов Иван Иванович</w:t><" .
            "/w:r><w:r w:rsidR=\"00FE09C8\" w:rsidRPr=\"00FE09C8\"><w:t xml:space=\"preserve\">, </w:t></w:r><w:r" .
            " w:rsidR=\"00FE09C8\"><w:t>возраст</w:t></w:r><w:r w:rsidR=\"00FE09C8\" w:rsidRPr=\"00FE09C8\"><w:t>" .
            ":</w:t></w:r><w:r w:rsidR=\"005E422E\"><w:t xml:space=\"preserve\"> </w:t></w:r><w:r w:rsidR=\"005E4" .
            "22E\" w:rsidRPr=\"00DE2B94\"><w:t>50</w:t></w:r></w:p><w:p w14:paraId=\"78D1FA3D\" w14:textId=\"55EF" .
            "7A47\" w:rsidR=\"00AF4B5B\" w:rsidRDefault=\"00DE2B94\" w:rsidP=\"0031223E\"><w:pPr><w:pStyle w:val=" .
            "\"a4\"/></w:pPr><w:r><w:t xml:space=\"preserve\">- </w:t></w:r><w:proofErr w:type=\"gramStart\"/><w:" .
            "r w:rsidRPr=\"00DE2B94\"><w:t xml:space=\"preserve\">Иванова Алена Ивановна</w:t></w:r><w:r w:rsidR=" .
            "\"00126257\" w:rsidRPr=\"00126257\"><w:t xml:space=\"preserve\"> </w:t></w:r></w:p><w:p w14:paraId=" .
            "\"78D1FA3D\" w14:textId=\"55EF7A47\" w:rsidR=\"00AF4B5B\" w:rsidRDefault=\"00DE2B94\" w:rsidP=\"00312" .
            "23E\"><w:pPr><w:pStyle w:val=\"a4\"/></w:pPr><w:r><w:t xml:space=\"preserve\">- </w:t></w:r><w:proofErr " .
            "w:type=\"gramStart\"/><w:r w:rsidRPr=\"00DE2B94\"><w:t xml:space=\"preserve\">Иванов Михаил Иванович" .
            "</w:t></w:r><w:r w:rsidR=\"00126257\" w:rsidRPr=\"00126257\"><w:t xml:space=\"preserve\"> </w:t></w:" .
            "r></w:p><w:p w14:paraId=\"45D5C362\" w14:textId=\"79C82D83\" w:rsidR=\"00AF4B5B\" w:rsidRDefault=\"0" .
            "05D75A1\" w:rsidP=\"006B0D50\"><w:pPr><w:pStyle w:val=\"a4\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId" .
            " w:val=\"2\"/></w:numPr></w:pPr><w:proofErr w:type=\"gramStart\"/><w:r w:rsidRPr=\"005D75A1\"><w:t x" .
            "ml:space=\"preserve\">Петров Петр Петрович</w:t></w:r><w:r w:rsidR=\"00FE09C8\" w:rsidRPr=\"00FE09C8" .
            "\"><w:t xml:space=\"preserve\">, </w:t></w:r><w:r w:rsidR=\"00FE09C8\"><w:t>возраст</w:t></w:r><w:r " .
            "w:rsidR=\"00FE09C8\" w:rsidRPr=\"00FE09C8\"><w:t>:</w:t></w:r><w:r w:rsidR=\"005E422E\"><w:t xml:spa" .
            "ce=\"preserve\"> </w:t></w:r><w:r w:rsidR=\"005E422E\" w:rsidRPr=\"00DE2B94\"><w:t>30</w:t></w:r></w" .
            ":p><w:p w14:paraId=\"4B66446E\" w14:textId=\"2D5C0B86\" w:rsidR=\"0024376E\" w:rsidRDefault=\"002437" .
            "6E\" w:rsidP=\"00AF4B5B\"/><w:p w14:paraId=\"3F32EC2C\" w14:textId=\"77777777\" w:rsidR=\"0024376E\"" .
            " w:rsidRPr=\"008D21C0\" w:rsidRDefault=\"0024376E\" w:rsidP=\"00AF4B5B\"/><w:sectPr w:rsidR=\"002437" .
            "6E\" w:rsidRPr=\"008D21C0\"><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"8" .
            "50\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/><w:cols w:s" .
            "pace=\"708\"/><w:docGrid w:linePitch=\"360\"/></w:sectPr></w:body></w:document>\n"
        );
        //$reporter->save("./tmp/report4.docx");
    }
}
