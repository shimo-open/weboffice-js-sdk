/**
 * SDK capability method identifiers.
 *
 * Values are public SDK method paths. Keeping the object tree aligned with the
 * public facade makes the accepted canIUse arguments discoverable without
 * exposing product-specific API namespaces.
 */
export const OfficeSDKMethods = {
  ActiveOutline: {
    Editor: {
      GetEditMode: 'ActiveOutline.Editor.GetEditMode',
      Document: {
        GetContent: 'ActiveOutline.Editor.Document.GetContent',
        GetTitleContent: 'ActiveOutline.Editor.Document.GetTitleContent',
        SetTitleContent: 'ActiveOutline.Editor.Document.SetTitleContent',
        Font: {
          SetTextColor: 'ActiveOutline.Editor.Document.Font.SetTextColor',
          SetHighLightColor:
            'ActiveOutline.Editor.Document.Font.SetHighLightColor',
          SetBold: 'ActiveOutline.Editor.Document.Font.SetBold',
          SetItalic: 'ActiveOutline.Editor.Document.Font.SetItalic',
          SetUnderline: 'ActiveOutline.Editor.Document.Font.SetUnderline',
          SetStrike: 'ActiveOutline.Editor.Document.Font.SetStrike'
        },
        Markdown: {
          GetMarkdown: 'ActiveOutline.Editor.Document.Markdown.GetMarkdown',
          AppendMarkdown:
            'ActiveOutline.Editor.Document.Markdown.AppendMarkdown',
          InsertMarkdown:
            'ActiveOutline.Editor.Document.Markdown.InsertMarkdown',
          ValidateMarkdown:
            'ActiveOutline.Editor.Document.Markdown.ValidateMarkdown'
        },
        Content: {
          ReplaceSelection:
            'ActiveOutline.Editor.Document.Content.ReplaceSelection',
          ReplaceAllContent:
            'ActiveOutline.Editor.Document.Content.ReplaceAllContent'
        }
      }
    },
    Service: {
      User: {
        GetUserInfo: 'ActiveOutline.Service.User.GetUserInfo'
      },
      Permission: {
        GetDocumentPermission:
          'ActiveOutline.Service.Permission.GetDocumentPermission'
      },
      Export: {
        DownloadDocument: 'ActiveOutline.Service.Export.DownloadDocument'
      },
      Collaboration: {
        GetSaveStatus: 'ActiveOutline.Service.Collaboration.GetSaveStatus'
      }
    },
    Env: {
      Language: {
        GetLanguage: 'ActiveOutline.Env.Language.GetLanguage'
      },
      DocsMode: {
        GetDocsMode: 'ActiveOutline.Env.DocsMode.GetDocsMode'
      }
    },
    Sub: {
      OnDocumentChange: 'ActiveOutline.Sub.OnDocumentChange'
    }
  }
} as const

type LeafValues<T> = T extends string
  ? T
  : T extends object
  ? {
      [K in keyof T]: LeafValues<T[K]>
    }[keyof T]
  : never

export type OfficeSDKMethodPath = LeafValues<typeof OfficeSDKMethods>
