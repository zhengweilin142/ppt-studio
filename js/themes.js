/* ============================================================
 * PPT Studio — 主题系统
 * 每个主题是一组颜色 + 字体，同时驱动网页预览与 PPTX 导出。
 * 全部为经典脚本（无 ESM），因此可以直接 file:// 双击打开。
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = (G.PPT = G.PPT || {});

  /**
   * 主题字段说明：
   *  bg        幻灯片主背景
   *  bgAlt     渐变/封面第二个色
   *  surface   卡片、圆角块底色
   *  border    描边色
   *  text      主文字
   *  muted     次要文字
   *  accent    强调色（1）
   *  accent2   强调色（2）
   *  dark      是否深色主题（影响预览阴影与默认假设）
   *  latin     西文字体
   *  ea        中日韩字体
   */
  var themes = {
    ocean: {
      id: 'ocean', name: '深海商务',
      bg: 'FFFFFF', bgAlt: 'EEF4FF', surface: 'F5F8FF', border: 'D6E2F8',
      text: '0F1F3D', muted: '5A6B8C', accent: '2563EB', accent2: '0EA5E9',
      dark: false, latin: 'Segoe UI', ea: 'Microsoft YaHei'
    },
    midnight: {
      id: 'midnight', name: '暗夜科技',
      bg: '0B1220', bgAlt: '132043', surface: '16233F', border: '2A3C63',
      text: 'F2F6FF', muted: '96A7C8', accent: '38BDF8', accent2: '818CF8',
      dark: true, latin: 'Segoe UI', ea: 'Microsoft YaHei'
    },
    ink: {
      id: 'ink', name: '极简墨白',
      bg: 'FFFFFF', bgAlt: 'F4F4F5', surface: 'F4F4F5', border: 'E1E1E4',
      text: '16161A', muted: '74747E', accent: '16161A', accent2: 'E0443E',
      dark: false, latin: 'Segoe UI', ea: 'Microsoft YaHei'
    },
    forest: {
      id: 'forest', name: '森野绿',
      bg: 'FBFDF9', bgAlt: 'E9F5E4', surface: 'EFF6EC', border: 'D2E5CA',
      text: '10261A', muted: '4E6B58', accent: '15A34A', accent2: '84CC16',
      dark: false, latin: 'Segoe UI', ea: 'Microsoft YaHei'
    },
    sunset: {
      id: 'sunset', name: '落日渐橙',
      bg: 'FFFBF7', bgAlt: 'FFEEDF', surface: 'FFF2E8', border: 'FBD8BC',
      text: '2B1608', muted: '8A5A3B', accent: 'EA580C', accent2: 'F59E0B',
      dark: false, latin: 'Segoe UI', ea: 'Microsoft YaHei'
    },
    violet: {
      id: 'violet', name: '紫罗兰',
      bg: 'FAF8FF', bgAlt: 'EFE8FF', surface: 'F2EDFF', border: 'DCD2FB',
      text: '1E1233', muted: '6B5B95', accent: '7C3AED', accent2: 'EC4899',
      dark: false, latin: 'Segoe UI', ea: 'Microsoft YaHei'
    }
  };

  PPT.themes = themes;
  PPT.themeList = ['ocean', 'midnight', 'ink', 'forest', 'sunset', 'violet'].map(function (k) {
    return { id: themes[k].id, name: themes[k].name, accent: themes[k].accent, bg: themes[k].bg, dark: themes[k].dark };
  });
  PPT.getTheme = function (id) { return themes[id] || themes.ocean; };

  if (typeof module !== 'undefined' && module.exports) module.exports = PPT;
})(typeof window !== 'undefined' ? window : globalThis);
