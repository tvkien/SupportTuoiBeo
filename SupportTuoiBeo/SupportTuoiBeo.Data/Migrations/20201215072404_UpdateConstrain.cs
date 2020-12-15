using Microsoft.EntityFrameworkCore.Migrations;

namespace SupportTuoiBeo.Data.Migrations
{
    public partial class UpdateConstrain : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_UserDetails_MaKH",
                table: "UserDetails");

            migrationBuilder.AlterColumn<string>(
                name: "MaKH",
                table: "UserDetails",
                type: "nvarchar(max)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(450)",
                oldNullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "MaKH",
                table: "UserDetails",
                type: "nvarchar(450)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(max)",
                oldNullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_UserDetails_MaKH",
                table: "UserDetails",
                column: "MaKH",
                unique: true,
                filter: "[MaKH] IS NOT NULL");
        }
    }
}
