import os
import json
import clr
import smartlayout


clr.AddReference('pptxlib')     # 自作のC#クラスライブラリ
from pptxlib import PptxDrawer # C#クラス
from pptxlib import PptxParser # C#クラス

def json2pptx(json_fiilename, pptx_filename):
    try:
        doc = smartlayout.Document(json_fiilename)
        pptxdrawer = PptxDrawer()

        for page in doc.pages:
            page_str = page.dumps()
            slide_id = pptxdrawer.add_slide(page_str)

            for kogumi in page.kogumis:
                shape_list = []
                for obj in kogumi.document.pages[0].objects:
                    obj_str = obj.dump()
                    shape_id = pptxdrawer.add_shape(slide_id, obj_str)
                    shape_list.append(shape_id)
                group_id = pptxdrawer.add_group(slide_id, shape_list)

            for obj in page.objects:
                obj_str = obj.dump()
                shape_id = pptxdrawer.add_shape(slide_id, obj_str)

        pptxdrawer.save(pptx_filename)
    except Exception as e:
        print("ERROR except in json2pptx() : ", e)
    return


def pptx2json(pptx_filename, json_fiilename):
    try:
        pptxparser = PptxParser(pptx_filename)
        doc = smartlayout.Document()

        slide_id_list = pptxparser.get_slide_id_list()
        for slide_id in slide_id_list:
            page = smartlayout.Page()

            kogumi_id_list = pptxparser.get_kogumi_id_list(slide_id)
            for kogumi_id in kogumi_id_list:
                kogumi = smartlayout.Kogumi()
                obj_id_list = pptxparser.get_kogumi_obj_id_list(slide_id, kogumi_id)
                for obj_id in obj_id_list:
                    obj_str = pptxparser.get_obj_in_kogumi(slide_id, kogumi_id, obj_id)
                    obj = create_object(obj_str)
                    kogumi.append_object(obj)
                page.append_kogumi(kogumi)

            obj_id_list = pptxparser.get_obj_id_list(slide_id)
            for obj_id in obj_id_list:
                obj_str = pptxparser.get_obj(slide_id, obj_id)
                obj = create_object(obj_str)
                page.append_object(obj)

            doc.append_page(page)
    except Exception as e:
        print("ERROR except in json2pptx() : ", e)
    return

def create_object(obj_str):
    obj_json = json.loads(obj_str)
    obj = None
    if obj_json["type"] == "text":
        obj = smartlayout.Text()
        obj.attribute = obj_json["attribute"]
        obj.x = obj_json["x"]
        obj.y = obj_json["y"]
        obj.width = obj_json["width"]
        obj.height = obj_json["height"]
        obj.text = obj_json["text"]
    elif obj_json["type"] == "image":
        obj = smartlayout.Image()
    elif obj_json["type"] == "Rect":
        obj = smartlayout.Image()
    elif obj_json["type"] == "circle":
        obj = smartlayout.Image()
    return obj