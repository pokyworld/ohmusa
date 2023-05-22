<% %>
  <table class="cart" cellspacing="0">
    <tbody>
<%
  For i = 1 To 5
%>      
      <tr class="cart-item">
        <td>
          <a class="delete-button" href="#" data-item="403" title="Remove this Item">
            <img src="../../images/delete.gif" alt="Remove Item">
          </a>
        </td>
        <td class="image-wrapper">
          <a href="#" class="product-image" data-product="2158" title="View Product Details">
            <img src="../../thumbimages/2158.jpg" alt="View Product">
          </a>
        </td>
        <td style="padding:0;">
          <table class="order-details" cellspacing="0">
            <tr class="header">
              <th class="left">SKU</th>
              <th colspan="3" class="left">Name</th>
              <th class="center">Stock</th>
              <th class="center">Quantity</th>
            </tr>
            <tr>
              <td class="left">AJ011</td>
              <td colspan="3" class="left">1829 Yellow Stephenson Rocket Steam Locomotive</td>
              <td class="center">On Order</td><!-- In Stock vs On order -->
              <td>
                <table class="quantity-controls" cellspacing="0">
                  <tr>
                    <td><a href="#" class="decrement" title="Remove 1 item"><button>-</button></a></td>
                    <td>
                        <span class="value">1</span>
                        <input class="value" type="hidden" value="1" />
                    </td>
                    <td class="center"><a href="#" class="increment" title="Add 1 extra item"><button>+</button></a>
                    </td>
                  </tr>
                </table>
              </td>

            </tr>
          </table>
          <table class="order-price" cellspacing="0">
            <tr class="header">
              <td colspan="3">&nbsp;</td>
              <th class="right">Price</th>
              <th class="right">Tax</th>
              <th class="right">SubTotal</th>
            </tr>
            <tr>
              <td colspan="3" class="right"><small>** Shipments to CA, US taxable at <span class="tax-rate">8.75</span>%</small></td>
              <td class="right">$&nbsp;<span class="price">172.92</span></td>
              <td class="right">$&nbsp;<span class="tax">15.13</span></td>
              <td class="right">$&nbsp;<span class="subtotal">188.05</span></td>
            </tr>
          </table>
        </td>
      </tr>
<%
  Next
%>
      <!-- </div> -->
    </tbody>
  </table>