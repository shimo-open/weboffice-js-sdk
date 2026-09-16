import type {
  HeaderBarsAnchorRect,
  HeaderBarsSourceAnchorRect
} from './OfficeSDK.headerBars'

/** iframe 内容区矩形，仅取换算需要的字段。 */
export interface HeaderBarsIframeRect {
  left: number
  top: number
  width: number
  height: number
}

/** 宿主视口尺寸。 */
export interface HeaderBarsViewportSize {
  width: number
  height: number
}

/**
 * 将 iframe 视口坐标系下的锚点矩形换算到宿主页面坐标系。
 *
 * 输入：
 *  - anchorRect：iframe 内部产生的矩形，viewport 为来源视口尺寸；
 *  - iframeRect：编辑器 iframe 元素在宿主页面的矩形，缺省时按无缩放、无偏移处理；
 *  - viewport：宿主视口尺寸。
 * 输出：可序列化的宿主坐标系矩形，coordinateSpace 固定为 callback-window-viewport。
 *
 * 缩放比例按来源视口尺寸计算，因此 iframe 被 CSS 缩放时仍能得到正确位置；
 * 不承诺旋转、倾斜或 3D transform 场景。
 */
export function toHeaderBarsCallbackAnchorRect(
  anchorRect: HeaderBarsSourceAnchorRect,
  iframeRect: HeaderBarsIframeRect | null | undefined,
  viewport: HeaderBarsViewportSize
): HeaderBarsAnchorRect {
  const sourceViewportWidth = anchorRect.viewport.width || 1
  const sourceViewportHeight = anchorRect.viewport.height || 1
  const scaleX =
    (iframeRect?.width ?? sourceViewportWidth) / sourceViewportWidth
  const scaleY =
    (iframeRect?.height ?? sourceViewportHeight) / sourceViewportHeight
  const offsetX = iframeRect?.left ?? 0
  const offsetY = iframeRect?.top ?? 0
  const left = offsetX + anchorRect.left * scaleX
  const top = offsetY + anchorRect.top * scaleY
  const right = offsetX + anchorRect.right * scaleX
  const bottom = offsetY + anchorRect.bottom * scaleY
  return {
    left,
    top,
    right,
    bottom,
    width: right - left,
    height: bottom - top,
    viewport: {
      width: viewport.width,
      height: viewport.height
    },
    coordinateSpace: 'callback-window-viewport'
  }
}
