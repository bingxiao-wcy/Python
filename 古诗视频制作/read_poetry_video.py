import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import cv2
import base64
import requests
import json
from moviepy.editor import concatenate_videoclips, VideoFileClip, AudioFileClip, ImageClip
import time
from pypinyin import pinyin
import shutil
import random
# ========================================
# === 用户可配置参数 ======================
# ========================================

# API配置
API_KEY = "jV3z2EK0O2Y5qBvvpFd4ZIgj"    # 百度API Key
SECRET_KEY = "SlBGKaLSx2SI3GA3fjzeSzurAu1SfC2b"  # 百度Secret Key
TTS_URL = "http://tsn.baidu.com/text2audio"

# 文件路径配置
INPUT_EXCEL_PATH = 'file/text.xlsx'      # 输入Excel文件路径
BACKGROUND_IMAGE_BASE_DIR = 'file'       # 背景图片基础文件夹
OUTPUT_DIR = 'result'                    # 输出文件夹
COVER_DIR = '封面'

# 视频参数
VIDEO_FPS = 24
VIDEO_CODEC = 'libx264'
AUDIO_CODEC = 'aac'

# 文字样式参数
CHINESE_FONT_PATH = "simkai.ttf"    # 楷体字体路径
PINYIN_FONT_PATH = "simhei.ttf"     # 黑体字体路径
ENGLISH_FONT_PATH = "arial.ttf"     # 英文字体路径

CHINESE_FONT_SIZE = 60
PINYIN_FONT_SIZE = 35
ENGLISH_FONT_SIZE = 40

TEXT_COLOR = (0, 0, 0)              # 黑色
HIGHLIGHT_COLOR = (255, 0, 0)       # 红色

CHAR_SPACING = 40                   # 字符间距
PINYIN_OFFSET = 40                  # 拼音偏移量
LINE_HEIGHT = 180                   # 行高

# 处理的工作表列表及对应配置
SHEET_CONFIG = [
    # {
    #     'name': '春晓',
    #     'intro_english': 'Please enjoy the Chinese ancient poem ',
    #     'intro_chinese': '接下来是中文版 '
    # },
    # {
    #     'name': '望庐山瀑布',
    #     'intro_english': 'Please enjoy the Chinese ancient poem ',
    #     'intro_chinese': '接下来是中文版 '
    # },
    {
        'name': '清明',
        'intro_english': 'Please enjoy the Chinese ancient poem ',
        'intro_chinese': '接下来是中文版 '
    },
    {
        'name': '月下独酌',
        'intro_english': 'Please enjoy the Chinese ancient poem ',
        'intro_chinese': '接下来是中文版 '
    },
    {
        'name': '春夜喜雨',
        'intro_english': 'Please enjoy the Chinese ancient poem ',
        'intro_chinese': '接下来是中文版 '
    }
]

# ========================================
# === 代码主体 ============================
# ========================================

def main():

    # 处理每个工作表配置
    for config in SHEET_CONFIG:
        sheet_name = config['name']
        sheet_dir = os.path.join(OUTPUT_DIR, sheet_name)
        os.makedirs(sheet_dir, exist_ok=True)
        clean_old_files(sheet_dir)

        # 创建输出目录
        sheet_output_image_dir = os.path.join(sheet_dir, 'output_images')
        sheet_output_video_dir = os.path.join(sheet_dir, 'output_video')
        sheet_output_cover_dir = os.path.join(sheet_dir, 'cover')
        os.makedirs(sheet_output_image_dir, exist_ok=True)
        os.makedirs(sheet_output_video_dir, exist_ok=True)
        # os.makedirs(sheet_output_cover_dir, exist_ok=True)
        
        # 清理旧文件
        clean_old_files(sheet_output_image_dir)
        clean_old_files(sheet_output_video_dir)
        # clean_old_files(sheet_output_cover_dir)

        # # 背景图片路径根据工作表名称拼接
        background_image_dir = os.path.join(BACKGROUND_IMAGE_BASE_DIR, sheet_name)
        # create_cover(INPUT_EXCEL_PATH, COVER_DIR, sheet_output_cover_dir, sheet_name)
        # 生成图片
        image_files = generate_images(INPUT_EXCEL_PATH, background_image_dir, sheet_output_image_dir, sheet_name)
        
        if not image_files:
            print(f"未生成图片（工作表：{sheet_name}），跳过此工作表")
            continue
        
        # 生成视频
        generate_sheet_video(INPUT_EXCEL_PATH, image_files, sheet_output_video_dir, sheet_name, config)
    
    # 清理临时文件
    clean_temp_files(OUTPUT_DIR)
    
    print(f"所有操作完成！")

def create_cover(excel_path, background_dir, output_dir, sheet_name):
    # 随机选择一张背景图片
    background_files = [f for f in os.listdir(background_dir) if f.endswith(('.png', '.jpg', '.jpeg'))]   
    background_file = random.choice(background_files)
    source_path = os.path.join(background_dir, background_file)
    destination_path = os.path.join(output_dir, background_file)
    shutil.copy2(source_path, destination_path)
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    # 定义文字内容
    chinese_text = df.iloc[0]['Chinese'].strip()  # 假设取第一行的中文内容
    english_text = df.iloc[0]['English'].strip()
    new_data = {
    'Chinese': [chinese_text],
    'English': [english_text]
    }
    new_df = pd.DataFrame(new_data)
    text_color = TEXT_COLOR
    alpha = 230
    generate_images(new_df, output_dir, output_dir, sheet_name, text_color, alpha)


def clean_old_files(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'删除文件 {file_path} 失败：{e}')

def clean_temp_files(directory):
    temp_extensions = ['.mp3', '_temp.mp4']
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if any(filename.endswith(ext) for ext in temp_extensions):
                file_path = os.path.join(root, filename)
                try:
                    os.remove(file_path)
                    print(f"删除临时文件：{file_path}")
                except Exception as e:
                    print(f'删除临时文件 {file_path} 失败：{e}')

# 定义函数：将中文转换为拼音（带声调）
def chinese_to_pinyin(text):
    pinyin_list = pinyin(text, style='TONE')
    pinyin_str = " ".join([item[0] for item in pinyin_list])
    # 删除拼音结尾的标点符号（如果有）
    pinyin_str = pinyin_str.rstrip(",.!?;")
    return pinyin_str

def generate_images(excel_path, background_dir, output_dir, sheet_name, text_color = None, alpha = 200):
    # 读取Excel数据
    try:
        if isinstance(excel_path, pd.DataFrame):
            # 如果是 DataFrame，则直接复制到目标路径
            df = excel_path
        elif isinstance(excel_path, str):
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
        else:
            print("错误：传入的参数必须是 DataFrame 或字符串路径")
    except ValueError:
        print(f"工作表 {sheet_name} 不存在，跳过")
        return []
    
    # 加载背景图片
    background_files = [f for f in os.listdir(background_dir) if f.endswith('.png')]
    background_count = len(background_files)
    
    if background_count == 0:
        print("未找到背景图片，程序退出")
        return []
    
    # 设置字体
    chinese_font = ImageFont.truetype(CHINESE_FONT_PATH, CHINESE_FONT_SIZE)
    pinyin_font = ImageFont.truetype(PINYIN_FONT_PATH, PINYIN_FONT_SIZE)
    english_font = ImageFont.truetype(ENGLISH_FONT_PATH, ENGLISH_FONT_SIZE)
    
    image_files = []
    
    for i in range(len(df)):
    # 加载背景图片
        background_index = i % background_count
        background_path = os.path.join(background_dir, background_files[background_index])
        background_image = Image.open(background_path)
        image_width, image_height = background_image.size
        
        # 裁剪背景图片
        background_image = background_image.crop((0, 0, image_width, image_height - 70))
        image_width, image_height = background_image.size  # 更新尺寸
        
        # 创建一个带有半透明文字区域的图像
        image = background_image.copy().convert('RGBA')
        overlay = Image.new('RGBA', (image_width, image_height), (0, 0, 0 ,0))
        draw_overlay = ImageDraw.Draw(overlay)
        draw = ImageDraw.Draw(image)
        
        # 计算最长文本行的宽度，以确定透明层的宽度
        max_width = 0
        for index, row in df.iterrows():
            chinese_text = row['Chinese']
            pinyin_text = chinese_to_pinyin(chinese_text)
            english_text = row['English']
            
            # 计算每行文本的宽度
            chinese_text_width = draw.textlength(chinese_text, font=chinese_font)
            pinyin_text_width = draw.textlength(pinyin_text, font=pinyin_font)
            english_text_width = draw.textlength(english_text, font=english_font)
            
            # 找出最长的文本行宽度
            current_max = max(chinese_text_width, pinyin_text_width, english_text_width)
            if current_max > max_width:
                max_width = current_max
    
        # 设置透明层宽度为最长文本行宽度左右各加5像素
        overlay_width = max_width + 80
        overlay_height = len(df) * LINE_HEIGHT + LINE_HEIGHT # 每行高度为120像素
        overlay_x = (image_width - overlay_width) // 2
        overlay_y = (image_height - overlay_height) // 2
        
        # 绘制半透明文字区域
        draw_overlay.rectangle(
            [overlay_x, overlay_y, overlay_x + overlay_width, overlay_y + overlay_height],
            fill=(255, 255, 255, alpha)  # 半透明白色背景
        )
        
        # 合并背景图片和半透明文字区域
        image = Image.alpha_composite(image, overlay).convert('RGB')
        draw = ImageDraw.Draw(image)
        
        # 计算文本在图片上的起始位置，使其居中
        total_text_height = len(df) * LINE_HEIGHT - LINE_HEIGHT / 2
        top_offset = (image_height - total_text_height) // 2
        
        # 遍历每一行文本
        for index, row in df.iterrows():
            # 提取中文内容
            chinese_text = row['Chinese'].rstrip("，。,.!?; ")
            
            # 生成拼音
            pinyin_list = pinyin(chinese_text, style='TONE')  # 逐个字符生成拼音
            pinyin_text = [item[0].rstrip("，。,.!?; ") for item in pinyin_list]
            
            # 英文内容
            english_text = row['English']
            
            # 计算文本在图片上的位置，使其居中
            current_top = top_offset + index * LINE_HEIGHT
            
            # 计算中文文本宽度，以居中显示
            chinese_text_width = draw.textlength(chinese_text, font=chinese_font)
            # 调整计算方式：总宽度 = 中文文本宽度 + (字符数量 - 1) * 字符间距
            total_chinese_width = chinese_text_width + (len(chinese_text) - 1) * CHAR_SPACING
            chinese_text_x = (image_width - total_chinese_width) // 2  # 居中位置

            # 写入中文和拼音（设置字符间距）
            current_char_x = chinese_text_x

            # 确定颜色
            color = text_color if text_color is not None else (HIGHLIGHT_COLOR if index == i else TEXT_COLOR)

            for char_index, char in enumerate(chinese_text):
                # 获取当前字符的拼音
                current_pinyin = pinyin_text[char_index]

                # 计算当前字符的拼音宽度
                pinyin_width = draw.textlength(current_pinyin, font=pinyin_font)
                
                # 计算当前字符的宽度（包括字符本身和后面的间距）
                # 注意：最后一个字符不需要额外的间距
                if char_index < len(chinese_text) - 1:
                    char_width = draw.textlength(char, font=chinese_font) + CHAR_SPACING
                else:
                    char_width = draw.textlength(char, font=chinese_font)

                # 绘制拼音（居中对齐汉字）
                # 调整拼音位置：使拼音在汉字正上方居中
                pinyin_x = current_char_x + (char_width - pinyin_width) // 8
                draw.text(
                    (pinyin_x, current_top - PINYIN_OFFSET),  # 拼音位置在汉字上方
                    current_pinyin.rstrip("，。,.!?;"),
                    font=pinyin_font,
                    fill=color
                )
                
                # 绘制中文字符
                draw.text(
                    (current_char_x, current_top),
                    char,
                    font=chinese_font,
                    fill=color
                )
                
                # 更新x坐标，增加字符实际占用的宽度（包括字符本身和间距）
                current_char_x += char_width

            # 写入英文（放在中文下方，居中显示）
            english_text = english_text.rstrip("，。,.!?;")
            english_text_width = draw.textlength(english_text, font=english_font)
            english_text_x = (image_width - english_text_width) // 2
            draw.text(
                (english_text_x, current_top + 60),
                english_text,
                font=english_font,
                fill=color
            )
        
        # 保存图片（按行索引命名）
        output_path = os.path.join(output_dir, f'{sheet_name}_image_{i+1}.png')
        image.save(output_path)
        image_files.append(output_path)
        print(f"已生成图片: {output_path}")
    print(f"工作表 {sheet_name} - 所有图片生成完成！")
    return image_files

def get_token():
    token_url = f"https://openapi.baidu.com/oauth/2.0/token?grant_type=client_credentials&client_id={API_KEY}&client_secret={SECRET_KEY}"
    response = requests.get(token_url)
    return response.json().get("access_token")

def synthesize_speech(text, token, language="zh"):
    headers = {"Content-Type": "application/json"}
    data = {
        "tex": text,
        "lan": language,
        "tok": token,
        "ctp": 1,
        "cuid": "123456PYTHON",
        "spd": 5,
        "pit": 5,
        "vol": 5,
        "per": 4103
    }
    response = requests.post(TTS_URL, data=data, headers=headers)
    return response.content

def concatenate_with_crossfade(clips, duration=0.5):
    for i in range(len(clips) - 1):
        clips[i] = clips[i].crossfadeout(duration)
        clips[i+1] = clips[i+1].crossfadein(duration)
    return concatenate_videoclips(clips)

def generate_sheet_video(excel_path, image_files, output_dir, sheet_name, config):
    # 获取令牌
    token = get_token()
    if not token:
        print(f"工作表 {sheet_name} - 获取令牌失败，跳过")
        return
    
    # 读取文本内容
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except ValueError:
        print(f"工作表 {sheet_name} 不存在，跳过")
        return
    
    english_texts = df['English'].tolist()
    chinese_texts = df['Chinese'].tolist()

    # 生成介绍语音，添加工作表名称
    intro_english_text = config['intro_english']
    intro_chinese_text = config['intro_chinese']
    

    intro_english_audio = synthesize_speech(intro_english_text, token, language="en")
    intro_chinese_audio = synthesize_speech(intro_chinese_text, token, language="zh")
    
    # 保存介绍语音
    intro_english_path = os.path.join(output_dir, f'{sheet_name}_intro_english.mp3') if intro_english_audio else None
    intro_chinese_path = os.path.join(output_dir, f'{sheet_name}_intro_chinese.mp3') if intro_chinese_audio else None
    
    if intro_english_path:
        with open(intro_english_path, 'wb') as f:
            f.write(intro_english_audio)
    
    if intro_chinese_path:
        with open(intro_chinese_path, 'wb') as f:
            f.write(intro_chinese_audio)
    
    # 生成视频片段
    video_clips_path = []
    current_image_index = -1  # 重置当前显示的图片索引
    # 英文旁白部分
    if intro_english_path:
        try:
            audio_clip = AudioFileClip(intro_english_path)
            image_clip = ImageClip(image_files[0], duration=audio_clip.duration - 0.1)
            video_clip = image_clip.set_audio(audio_clip)
            video_path = os.path.join(output_dir, f'{sheet_name}_intro_english_temp.mp4')
            video_clip.write_videofile(video_path, codec=VIDEO_CODEC, audio_codec=AUDIO_CODEC, fps=VIDEO_FPS)
            print(f"English intro video segment generated: {video_path}")
            video_clips_path.append(video_path)
            audio_clip.close()
        except Exception as e:
            print(f"处理英文介绍时出错：{e}")
    # 英文
    for idx, text in enumerate(english_texts):
        audio_data = synthesize_speech(text, token, language="en")
        if not audio_data or b"err_msg" in audio_data:
            print(f"合成语音失败（英文内容：{text}），跳过")
            continue
        
        audio_path = os.path.join(output_dir, f'{sheet_name}_audio_english_{idx+1}.mp3')
        video_path = os.path.join(output_dir, f'{sheet_name}_video_english_{idx+1}_temp.mp4')
        with open(audio_path, 'wb') as f:
            f.write(audio_data)
        
        if not os.path.exists(audio_path):
            print(f"音频文件生成失败：{audio_path}，跳过")
            continue
        try:
            audio_clip = AudioFileClip(audio_path)
        except Exception as e:
            print(f"加载音频文件失败：{audio_path} - {e}，跳过")
            continue
        
        if not audio_clip:
            print(f"音频剪辑对象创建失败：{audio_path}，跳过")
            continue
        # 检查音频文件是否包含有效音频数据
        audio_clip = None
        try:
            audio_clip = AudioFileClip(audio_path)
        except Exception as e:
            print(f"Invalid audio file: {audio_path} - {e}")
            continue

        if audio_clip is None:
            print(f"Failed to load audio file: {audio_path}")
            continue
        audio_duration = audio_clip.duration
        print(f"Audio duration: {audio_duration} seconds")
        try:
            current_image_index = min(current_image_index + 1, len(image_files) - 1)
            current_image_path = image_files[current_image_index]
            print(current_image_index, current_image_path)
            image_clip = ImageClip(current_image_path, duration=audio_clip.duration - 0.1)
            video_clip = image_clip.set_audio(audio_clip)
            video_clip.write_videofile(video_path, codec=VIDEO_CODEC, audio_codec=AUDIO_CODEC, fps=VIDEO_FPS)
            print(f"English intro video segment generated: {video_path}")
            video_clips_path.append(video_path)
            audio_clip.close()
        except Exception as e:
            print(f"创建视频片段失败（图片：{image_files[idx]}，音频：{audio_path})- {e}，跳过")
    current_image_index = -1  # 重置当前显示的图片索引
    # 中文旁白部分
    if intro_chinese_path:
        try:
            audio_clip = AudioFileClip(intro_chinese_path)
            image_clip = ImageClip(image_files[0], duration=audio_clip.duration - 0.1)
            video_clip = image_clip.set_audio(audio_clip)
            video_path = os.path.join(output_dir, f'{sheet_name}_intro_chinese_temp.mp4')
            video_clip.write_videofile(video_path, codec=VIDEO_CODEC, audio_codec=AUDIO_CODEC, fps=VIDEO_FPS)
            print(f"English intro video segment generated: {video_path}")
            video_clips_path.append(video_path)
            audio_clip.close()
        except Exception as e:
            print(f"处理英文介绍时出错：{e}")

    # 中文
    for idx, text in enumerate(chinese_texts):
        audio_data = synthesize_speech(text, token, language="zh")
        if not audio_data or b"err_msg" in audio_data:
            print(f"合成语音失败（英文内容：{text}），跳过")
            continue
        
        audio_path = os.path.join(output_dir, f'{sheet_name}_audio_chinese_{idx+1}.mp3')
        video_path = os.path.join(output_dir, f'{sheet_name}_video_chinese_{idx+1}_temp.mp4')
        with open(audio_path, 'wb') as f:
            f.write(audio_data)
        
        if not os.path.exists(audio_path):
            print(f"音频文件生成失败：{audio_path}，跳过")
            continue
        try:
            audio_clip = AudioFileClip(audio_path)
        except Exception as e:
            print(f"加载音频文件失败：{audio_path} - {e}，跳过")
            continue
        
        if not audio_clip:
            print(f"音频剪辑对象创建失败：{audio_path}，跳过")
            continue
        # 检查音频文件是否包含有效音频数据
        audio_clip = None
        try:
            audio_clip = AudioFileClip(audio_path)
        except Exception as e:
            print(f"Invalid audio file: {audio_path} - {e}")
            continue

        if audio_clip is None:
            print(f"Failed to load audio file: {audio_path}")
            continue
        audio_duration = audio_clip.duration
        print(f"Audio duration: {audio_duration} seconds")
        try:
            current_image_index = min(current_image_index + 1, len(image_files) - 1)
            current_image_path = image_files[current_image_index]
            print(current_image_index, current_image_path)
            image_clip = ImageClip(current_image_path, duration=audio_clip.duration - 0.1)
            video_clip = image_clip.set_audio(audio_clip)
            video_clip.write_videofile(video_path, codec=VIDEO_CODEC, audio_codec=AUDIO_CODEC, fps=VIDEO_FPS)
            print(f"English intro video segment generated: {video_path}")
            video_clips_path.append(video_path)
            audio_clip.close()
        except Exception as e:
            print(f"创建视频片段失败（图片：{image_files[idx]}，音频：{audio_path})- {e}，跳过")

    if video_clips_path:
        video_clips = [VideoFileClip(video_file) for video_file in video_clips_path]
        final_clip = concatenate_with_crossfade(video_clips)
        final_path = os.path.join(output_dir, f'{sheet_name}_final.mp4')
        final_clip.write_videofile(final_path, codec=VIDEO_CODEC, audio_codec=AUDIO_CODEC, fps=VIDEO_FPS)
        print(f"工作表 {sheet_name} - 视频生成完成：{final_path}")
    else:
        print(f"工作表 {sheet_name} - 没有有效的视频片段可以合并")

if __name__ == "__main__":
    main()