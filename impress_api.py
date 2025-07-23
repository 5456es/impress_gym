import os


import uno
from com.sun.star.awt import Point, Size
from com.sun.star.beans import PropertyValue
from flask import Flask, jsonify, request
import logging
import sys

# 设置日志
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)
desktop = None
ctx=None


def extract_table_info(table_shape):
    """
    给定 com.sun.star.drawing.TableShape → 返回行数、列数与全部单元格内容
    {
        "rows": <int>,
        "columns": <int>,
        "data": [  # data[row][col] = str
            ["A1", "B1", ...],
            ["A2", "B2", ...],
            ...
        ]
    }
    """
    try:
        model = getattr(table_shape, "Model", None)
        if model is None:
            return {"error": "TableShape has no Model"}

        rows = model.Rows.getCount()          # 行数
        cols = model.Columns.getCount()       # 列数

        def _cell_to_str(cell):
            """兼容 getString / String / getText().getString() 三种写法"""
            if hasattr(cell, "getString"):
                return cell.getString()
            if hasattr(cell, "String"):
                return cell.String
            if hasattr(cell, "getText"):
                return cell.getText().getString()
            return ""

        data = []
        for r in range(rows):
            row_vals = []
            for c in range(cols):
                try:
                    # 必须走 model.getCellByPosition(col, row) :contentReference[oaicite:0]{index=0}
                    cell = model.getCellByPosition(c, r)
                    row_vals.append(_cell_to_str(cell))
                except Exception:
                    row_vals.append("")
            data.append(row_vals)

        return {"rows": rows, "columns": cols, "data": data}

    except Exception as e:
        import traceback
        return {
            "error": f"extract_table_info failed: {e}",
            "traceback": traceback.format_exc()
        }

def extract_formatting(shape):
    """提取文本格式信息"""
    formatting = {}
    try:
        text_cursor = shape.createTextCursor()
        formatting = {
            "font": text_cursor.CharFontName if hasattr(text_cursor, 'CharFontName') else "",
            "font_size": float(text_cursor.CharHeight) if hasattr(text_cursor, 'CharHeight') else 0,
            "color": text_cursor.CharColor if hasattr(text_cursor, 'CharColor') else 0,
            "bold": text_cursor.CharWeight == 150.0 if hasattr(text_cursor, 'CharWeight') else False,
            "italic": text_cursor.CharPosture != 0 if hasattr(text_cursor, 'CharPosture') else False,
            "strikeout": text_cursor.CharStrikeout != 0 if hasattr(text_cursor, 'CharStrikeout') else False,
            "alignment": {
                0: "left",
                1: "right",
                2: "center",
                3: "justify"
            }.get(text_cursor.ParaAdjust, "unknown") if hasattr(text_cursor, 'ParaAdjust') else "unknown"
        }
    except Exception as e:
        formatting = {"error": f"formatting extraction failed: {str(e)}"}
    return formatting

# 添加错误处理
@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"Unhandled exception: {e}", exc_info=True)
    return jsonify({"error": str(e)}), 500

def connect_to_libreoffice():
    """连接到正在运行的LibreOffice实例或启动新实例"""
    global ctx
    try:
        logger.info("尝试连接到 LibreOffice...")
        local_context = uno.getComponentContext()
        logger.debug("获取本地上下文成功")
        
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        logger.debug("创建解析器成功")
        
        ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
        logger.info("成功连接到 LibreOffice!")
        
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        return desktop
    except Exception as e:
        logger.error(f"连接 LibreOffice 失败: {e}", exc_info=True)
        return None

def get_current_presentation():
    """获取当前活动演示文稿"""
    global desktop
    if not desktop:
        desktop = connect_to_libreoffice()
        
        if not desktop:
            return None
    
    doc = desktop.getCurrentComponent()
    # 检查是否是演示文稿
    if doc and doc.supportsService("com.sun.star.presentation.PresentationDocument"):
        return doc
    return None

def get_current_slide(doc):
    """获取当前选中的幻灯片"""
    if not doc:
        return None
    
    try:
        controller = doc.getCurrentController()
        # 获取当前页面
        current_page = controller.getCurrentPage()
        return current_page
    except Exception as e:
        print(f"Error getting current slide: {e}")
        return None

def get_current_selection(doc):
    """获取当前选中的对象（shape），并提取文本及格式属性（若有）"""
    try:
        from com.sun.star.view import XSelectionSupplier

        controller = doc.getCurrentController()
        if not hasattr(controller, 'getSelection') and not isinstance(controller, XSelectionSupplier):
            return {"error": "Controller does not support selection (XSelectionSupplier)"}

        selection = controller.getSelection()
        if not selection:
            return {"status": "empty", "message": "No selection"}

        def shape_info_from_shape(shape, index=None):
            text = shape.getString() if hasattr(shape, "getString") else ""
            info = {
                "type": shape.getShapeType(),
                "text": text,
                "position": {"x": shape.Position.X, "y": shape.Position.Y},
                "size": {"width": shape.Size.Width, "height": shape.Size.Height}
            }
            if index is not None:
                info["index"] = index
            if text:
                info["formatting"] = extract_formatting(shape)
            if shape.getShapeType() == "com.sun.star.drawing.TableShape":
                info["table"] = extract_table_info(shape)
            return info

        # 多个选中对象
        if hasattr(selection, 'getCount'):
            count = selection.getCount()
            shapes = [
                shape_info_from_shape(selection.getByIndex(i), index=i)
                for i in range(count)
            ]
            return {
                "status": "success",
                "selection_count": count,
                "shapes": shapes
            }

        # 单个对象
        elif hasattr(selection, 'getShapeType'):
            return {
                "status": "success",
                "selection_count": 1,
                "shape": shape_info_from_shape(selection)
            }

        else:
            return {"status": "unknown selection type"}

    except Exception as e:
        import traceback
        return {
            "error": f"getCurrentSelection failed: {str(e)}",
            "traceback": traceback.format_exc()
        }


def get_selected_text(doc):
    """
    doc  : 选填，XModel；若为 None，则用 ctx 去拿 current component
    """
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    if doc is None:
        doc = desktop.getCurrentComponent()
    if not doc:
        return {"error": "no-document"}

    ctrl = doc.getCurrentController()

    try:
        helper = smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx
        )
        helper.executeDispatch(ctrl.getFrame(), ".uno:Copy", "", 0, ())

        clip = smgr.createInstanceWithContext(
            "com.sun.star.datatransfer.clipboard.SystemClipboard", ctx
        )
        xfer = clip.getContents()

        for flav in xfer.getTransferDataFlavors():
            if flav.MimeType.lower().startswith("text/plain"):
                return {
                    "status": "ok-clipboard",
                    "text": str(xfer.getTransferData(flav))
                }

    except Exception as e:
        print("clipboard fallback failed:", e)

    return {"error": "no-text-selection"}

def get_slide_by_index(doc, index):
    """通过索引获取幻灯片"""
    if not doc:
        return None
    
    try:
        draw_pages = doc.getDrawPages()
        if index < 0 or index >= draw_pages.getCount():
            return None
        return draw_pages.getByIndex(index)
    except Exception as e:
        print(f"Error getting slide by index: {e}")
        return None

def get_presentation_info(doc):
    """获取演示文稿基本信息"""
    if not doc:
        return {"error": "No presentation available"}
    
    try:
        draw_pages = doc.getDrawPages()
        controller = doc.getCurrentController()
        current_page = controller.getCurrentPage()
        
        # 查找当前页面的索引
        current_index = -1
        for i in range(draw_pages.getCount()):
            if draw_pages.getByIndex(i) == current_page:
                current_index = i
                break
        
        info = {
            "total_slides": draw_pages.getCount(),
            "current_slide_index": current_index,
            "presentation_name": doc.getTitle() if hasattr(doc, 'getTitle') else "Untitled"
        }
        
        return info
    except Exception as e:
        return {"error": str(e)}

def get_slide_content(slide, include_formatting=True):
    """获取幻灯片内容"""

    if not slide:
        return {"error": "No slide provided"}
    
    try:
        shapes = []
        shape_count = slide.getCount()
        
        for i in range(shape_count):
            shape = slide.getByIndex(i)
            shape_info = {
                "index": i,
                "type": shape.getShapeType(),
                "position": {"x": shape.Position.X, "y": shape.Position.Y},
                "size": {"width": shape.Size.Width, "height": shape.Size.Height}
            }
            
            # 检查是否包含文本
            if hasattr(shape, 'getString'):
                text = shape.getString()
                shape_info["text"] = text
                
                if include_formatting and text:
                    # 获取文本格式信息
                    
                    shape_info["formatting"] = extract_formatting(shape)
                    
            #  如果是table，获取基本信息，包括行列以及内容
            if shape.getShapeType() == "com.sun.star.drawing.TableShape":
                shape_info["table"] = extract_table_info(shape)
            shapes.append(shape_info)
        
        return {
            "status": "success",
            "shape_count": shape_count,
            "shapes": shapes
        }
    except Exception as e:
        return {"error": str(e)}

def add_text_shape(slide, text, x=1000, y=1000, width=10000, height=2000, formatting=None):
    """在幻灯片上添加文本框"""
    if not slide:
        return {"error": "No slide provided"}
    
    try:
        # 创建文本框
        doc = slide.getModel()  # 获取文档模型
        shape = doc.createInstance("com.sun.star.drawing.TextShape")
        
        # 设置位置和大小
        shape.Position = Point(x, y)
        shape.Size = Size(width, height)
        
        # 添加到幻灯片
        slide.add(shape)
        
        # 设置文本
        shape.setString(text)
        
        # 应用格式化
        if formatting:
            text_cursor = shape.createTextCursor()
            apply_text_formatting(text_cursor, formatting)
        
        return {
            "status": "success",
            "message": "Text shape added",
            "shape_index": slide.getCount() - 1
        }
    except Exception as e:
        return {"error": str(e)}

def update_shape_text(slide, shape_index, new_text, formatting=None):
    """更新形状中的文本"""
    if not slide:
        return {"error": "No slide provided"}
    
    try:
        if shape_index < 0 or shape_index >= slide.getCount():
            return {"error": "Invalid shape index"}
        
        shape = slide.getByIndex(shape_index)
        
        # 检查是否支持文本
        if not hasattr(shape, 'setString'):
            return {"error": "Shape does not support text"}
        
        # 设置新文本
        shape.setString(new_text)
        
        # 应用格式化
        if formatting:
            text_cursor = shape.createTextCursor()
            apply_text_formatting(text_cursor, formatting)
        
        return {"status": "success", "message": "Shape text updated"}
    except Exception as e:
        return {"error": str(e)}

def apply_text_formatting(text_cursor, formatting):
    """应用文本格式化"""
    try:
        if "font" in formatting and hasattr(text_cursor, 'CharFontName'):
            text_cursor.CharFontName = formatting["font"]
        if "font_size" in formatting and hasattr(text_cursor, 'CharHeight'):
            text_cursor.CharHeight = float(formatting["font_size"])
        if "color" in formatting and hasattr(text_cursor, 'CharColor'):
            text_cursor.CharColor = int(formatting["color"])
        if "bold" in formatting and hasattr(text_cursor, 'CharWeight'):
            text_cursor.CharWeight = 150.0 if formatting["bold"] else 100.0
        if "italic" in formatting and hasattr(text_cursor, 'CharPosture'):
            text_cursor.CharPosture = 1 if formatting["italic"] else 0
    except Exception as e:
        print(f"Error applying formatting: {e}")

def add_new_slide(doc, position=-1):
    """添加新幻灯片"""
    if not doc:
        return {"error": "No presentation available"}
    
    try:
        draw_pages = doc.getDrawPages()
        
        # 如果position为-1，添加到末尾
        if position == -1:
            position = draw_pages.getCount()
        
        # 插入新页面
        new_page = draw_pages.insertNewByIndex(position)
        
        return {
            "status": "success",
            "message": "New slide added",
            "slide_index": position,
            "total_slides": draw_pages.getCount()
        }
    except Exception as e:
        return {"error": str(e)}

def delete_slide(doc, slide_index):
    """删除幻灯片"""
    if not doc:
        return {"error": "No presentation available"}
    
    try:
        draw_pages = doc.getDrawPages()
        
        if slide_index < 0 or slide_index >= draw_pages.getCount():
            return {"error": "Invalid slide index"}
        
        # 删除页面
        page = draw_pages.getByIndex(slide_index)
        draw_pages.remove(page)
        
        return {
            "status": "success",
            "message": "Slide deleted",
            "total_slides": draw_pages.getCount()
        }
    except Exception as e:
        return {"error": str(e)}

# API 端点

@app.route('/api/connect', methods=['POST'])
def api_connect():
    """API端点：连接到LibreOffice"""
    global desktop
    global ctx
    try:
        desktop = connect_to_libreoffice()
        if desktop:
            return jsonify({"status": "success", "message": "Connected to LibreOffice"})
        else:
            return jsonify({"status": "error", "message": "Failed to connect to LibreOffice"}), 500
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/presentation/info', methods=['GET'])
def api_get_presentation_info():
    """API端点：获取演示文稿信息"""
    doc = get_current_presentation()
    result = get_presentation_info(doc)
    return jsonify(result)

@app.route('/api/slide/current', methods=['GET'])
def api_get_current_slide():
    """API端点：获取当前幻灯片内容"""
    include_formatting = request.args.get('include_formatting', 'false').lower() == 'true'
    
    doc = get_current_presentation()
    if not doc:
        return jsonify({"error": "No presentation available"}), 404
    
    slide = get_current_slide(doc)
    if not slide:
        return jsonify({"error": "No current slide"}), 404
    
    result = get_slide_content(slide, include_formatting)
    return jsonify(result)

@app.route('/api/slide/<int:index>', methods=['GET'])
def api_get_slide_by_index(index):
    """API端点：通过索引获取幻灯片内容"""
    include_formatting = request.args.get('include_formatting', 'false').lower() == 'true'
    
    doc = get_current_presentation()
    if not doc:
        return jsonify({"error": "No presentation available"}), 404
    
    slide = get_slide_by_index(doc, index)
    if not slide:
        return jsonify({"error": f"Slide {index} not found"}), 404
    
    result = get_slide_content(slide, include_formatting)
    return jsonify(result)

@app.route('/api/slide/add-text', methods=['POST'])
def api_add_text_to_slide():
    """API端点：向幻灯片添加文本框"""
    data = request.get_json()
    text = data.get('text', '')
    slide_index = data.get('slide_index', None)  # None表示当前幻灯片
    x = data.get('x', 1000)
    y = data.get('y', 1000)
    width = data.get('width', 10000)
    height = data.get('height', 2000)
    formatting = data.get('formatting', None)
    
    if not text:
        return jsonify({"error": "Missing 'text' parameter"}), 400
    
    doc = get_current_presentation()
    if not doc:
        return jsonify({"error": "No presentation available"}), 404
    
    if slide_index is None:
        slide = get_current_slide(doc)
    else:
        slide = get_slide_by_index(doc, slide_index)
    
    if not slide:
        return jsonify({"error": "Slide not found"}), 404
    
    result = add_text_shape(slide, text, x, y, width, height, formatting)
    return jsonify(result)

@app.route('/api/slide/update-shape', methods=['PUT'])
def api_update_shape_text():
    """API端点：更新形状文本"""
    data = request.get_json()
    slide_index = data.get('slide_index', None)
    shape_index = data.get('shape_index')
    new_text = data.get('text', '')
    formatting = data.get('formatting', None)
    
    if shape_index is None:
        return jsonify({"error": "Missing 'shape_index' parameter"}), 400
    
    doc = get_current_presentation()
    if not doc:
        return jsonify({"error": "No presentation available"}), 404
    
    if slide_index is None:
        slide = get_current_slide(doc)
    else:
        slide = get_slide_by_index(doc, slide_index)
    
    if not slide:
        return jsonify({"error": "Slide not found"}), 404
    
    result = update_shape_text(slide, shape_index, new_text, formatting)
    return jsonify(result)

@app.route('/api/slide/selection', methods=['GET'])
def api_get_selection():
    """API端点：获取当前选中的对象"""
    doc = get_current_presentation()
    if not doc:
        return jsonify({"error": "No presentation available"}), 404
    
    result = get_current_selection(doc)
    return jsonify(result)

@app.route('/api/slide/text-selection', methods=['GET'])
def api_get_text_selection():
    doc = get_current_presentation()
    if not doc:
        return jsonify({"error": "No presentation available"}), 404

    result = get_selected_text(doc)
    return jsonify(result)

@app.route('/api/slide/new', methods=['POST'])
def api_add_slide():
    """API端点：添加新幻灯片"""
    data = request.get_json()
    position = data.get('position', -1)
    
    doc = get_current_presentation()
    result = add_new_slide(doc, position)
    return jsonify(result)

@app.route('/api/slide/<int:index>', methods=['DELETE'])
def api_delete_slide(index):
    """API端点：删除幻灯片"""
    doc = get_current_presentation()
    result = delete_slide(doc, index)
    return jsonify(result)

@app.route('/api/health', methods=['GET'])
def api_health():
    """API健康检查"""
    try:
        doc = get_current_presentation()
        if doc:
            status = "connected"
            message = "Connected to Impress presentation"
        else:
            status = "no_presentation"
            message = "LibreOffice is running but no Impress presentation is open"
        
        return jsonify({
            "status": status,
            "message": message,
            "service": "impress-api"
        })
    except Exception as e:
        logger.error(f"Health check error: {e}")
        return jsonify({
            "status": "error",
            "message": str(e),
            "service": "impress-api"
        }), 503

if __name__ == '__main__':
    logger.info("启动 LibreOffice Impress API 服务...")
    logger.info(f"Python 路径: {sys.path}")
    logger.info(f"Python 版本: {sys.version}")
    
    # 测试 UNO 导入
    try:
        import uno
        logger.info("UNO 模块导入成功")
    except ImportError as e:
        logger.error(f"无法导入 UNO 模块: {e}")
        logger.error("请确保使用 LibreOffice 的 Python 或正确设置 PYTHONPATH")
    
    app.run(host='0.0.0.0', port=5011, debug=True)