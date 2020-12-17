declare function typeOf(
  value: any
): (
  "string"|"object"|"number"|"array"|
  "undefined"|"boolean"|"regexp"|
  "date"|"null"|"nan"|"function"
)

declare module typeOf {
  function Undefined(type: any): boolean
  function Boolean(type: any): boolean
  function String(type: any): boolean
  function Function(type: any): boolean
  function Array(type: any): boolean
  function Object(type: any): boolean
  function Null(type: any): boolean
  function Number(type: any): boolean
  function Date(type: any): boolean
  function RegExp(type: any): boolean
  function NaN(type: any): boolean
}

export = typeOf